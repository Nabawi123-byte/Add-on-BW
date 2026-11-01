using System.Net;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;
using MasterBW.Core.Ai;
using MasterBW.Core.Util;
using MasterBW.Engine;
using Newtonsoft.Json;

// =====================================================================================
// MasterBW.Engine — server AI lokal (127.0.0.1 saja, tidak bisa diakses dari jaringan).
// Dijalankan otomatis oleh add-in Word/Excel. Bisa juga dijalankan manual untuk uji:
//   MasterBW.Engine.exe --test "Buat rumus total penjualan"
// =====================================================================================

const string Version = "1.0.0";
var settings = MasterBWSettings.Load();
string baseDir = AppContext.BaseDirectory;
string modelsDir = args.SkipWhile(a => a != "--models").Skip(1).FirstOrDefault()
                   ?? (Directory.Exists(Path.Combine(baseDir, "models")) ? Path.Combine(baseDir, "models") : AppPaths.ModelsDir);

using var mutex = new Mutex(true, "MasterBW.Engine.Singleton", out bool firstInstance);
if (!firstInstance && !args.Contains("--test"))
{
    Console.WriteLine("MasterBW.Engine sudah berjalan.");
    return;
}

NativeCheck.Run(baseDir);
var host = new ModelHost(settings, modelsDir);
var speechPath = Path.IsPathRooted(settings.SpeechModelFile) ? settings.SpeechModelFile : Path.Combine(modelsDir, settings.SpeechModelFile);
var speech = new SpeechHost(speechPath);

// ---- mode uji di konsol ----
if (args.Contains("--test"))
{
    var prompt = args.SkipWhile(a => a != "--test").Skip(1).FirstOrDefault() ?? "Perkenalkan dirimu dalam dua kalimat.";
    Console.OutputEncoding = Encoding.UTF8;
    Console.WriteLine("Model: " + host.ResolveModelPath());
    var res = await host.GenerateAsync(new GenerateRequest
    {
        Messages = { ChatMessage.System(Prompts.Persona), ChatMessage.User(prompt) },
        MaxTokens = 400,
        Reasoning = ReasoningMode.Cepat,
    }, c => Console.Write(c.Delta), CancellationToken.None);
    Console.WriteLine();
    Console.WriteLine(res.Error ?? $"[{res.CompletionTokens} token, {res.Seconds:0.0} dtk]");
    return;
}

// ---- pilih port bebas mulai dari pengaturan ----
int port = settings.EnginePort;
for (int i = 0; i < 20; i++, port++)
{
    try { var l = new TcpListener(IPAddress.Loopback, port); l.Start(); l.Stop(); break; }
    catch (SocketException) { }
}
var token = Convert.ToHexString(RandomNumberGenerator.GetBytes(16));

var builder = WebApplication.CreateSlimBuilder();
builder.Logging.ClearProviders();
builder.WebHost.ConfigureKestrel(o =>
{
    o.Listen(IPAddress.Loopback, port);
    o.Limits.MaxRequestBodySize = 64 * 1024 * 1024;
});
var app = builder.Build();
var lifetime = app.Lifetime;

// Semua endpoint (kecuali /health) wajib membawa token dari engine.json.
app.Use(async (ctx, next) =>
{
    if (ctx.Request.Path != "/health" && ctx.Request.Headers["X-MasterBW-Token"] != token)
    {
        ctx.Response.StatusCode = 401;
        await ctx.Response.WriteAsync("token tidak valid");
        return;
    }
    await next();
});

app.MapGet("/health", () => Results.Text(JsonConvert.SerializeObject(new EngineStatus
{
    Ok = true,
    Version = Version,
    ModelLoaded = host.Loaded,
    ModelName = host.Loaded ? host.ModelName : Path.GetFileNameWithoutExtension(host.ResolveModelPath()),
    ModelPath = host.ResolveModelPath(),
    ContextSize = host.ContextSize,
    Threads = host.Threads,
    SpeechAvailable = speech.Available,
    Busy = host.Busy,
    Error = NativeCheck.Error ?? (File.Exists(host.ResolveModelPath()) ? host.LoadError : "File model belum ada di " + modelsDir),
}), "application/json"));

app.MapPost("/v1/generate", async (HttpContext ctx) =>
{
    using var reader = new StreamReader(ctx.Request.Body, Encoding.UTF8);
    var body = await reader.ReadToEndAsync();
    var req = JsonConvert.DeserializeObject<GenerateRequest>(body) ?? new GenerateRequest();
    bool stream = ctx.Request.Query["stream"] == "1";
    var ct = ctx.RequestAborted;

    if (!stream)
    {
        var res = await host.GenerateAsync(req, null, ct);
        ctx.Response.ContentType = "application/json";
        await ctx.Response.WriteAsync(JsonConvert.SerializeObject(res), ct);
        return;
    }

    ctx.Response.ContentType = "text/event-stream";
    ctx.Response.Headers.CacheControl = "no-cache";
    var writeLock = new SemaphoreSlim(1, 1);
    async Task Send(StreamChunk c)
    {
        await writeLock.WaitAsync();
        try
        {
            await ctx.Response.WriteAsync("data: " + JsonConvert.SerializeObject(c) + "\n\n");
            await ctx.Response.Body.FlushAsync();
        }
        finally { writeLock.Release(); }
    }
    var pending = new List<Task>();
    var result = await host.GenerateAsync(req, c => pending.Add(Send(c)), ct);
    await Task.WhenAll(pending);
    await Send(new StreamChunk { Kind = "done", Result = result });
});

app.MapPost("/v1/transcribe", async (HttpContext ctx) =>
{
    var ms = new MemoryStream();
    await ctx.Request.Body.CopyToAsync(ms);
    ms.Position = 0;
    var resp = new TranscribeResponse();
    try { resp.Text = await speech.TranscribeAsync(ms, ctx.Request.Query["lang"].FirstOrDefault() ?? "id", ctx.RequestAborted); }
    catch (Exception ex) { resp.Error = ex.Message; Log.Error("Transkripsi gagal", ex); }
    return Results.Text(JsonConvert.SerializeObject(resp), "application/json");
});

app.MapPost("/v1/unload", async () => { await host.UnloadAsync(); return Results.Ok(); });
app.MapPost("/v1/shutdown", () => { lifetime.StopApplication(); return Results.Ok(); });

lifetime.ApplicationStarted.Register(() =>
{
    AppPaths.WriteJson(AppPaths.EngineEndpointFile, new EngineEndpoint { Port = port, Token = token, Pid = Environment.ProcessId });
    Log.Info($"MasterBW.Engine {Version} siap di 127.0.0.1:{port}, models={modelsDir}");
});
lifetime.ApplicationStopping.Register(() =>
{
    try { File.Delete(AppPaths.EngineEndpointFile); } catch { }
    host.Dispose();
    speech.Dispose();
});

// Matikan engine bila tidak ada Word/Excel yang berjalan (cek tiap 2 menit).
_ = Task.Run(async () =>
{
    if (!OperatingSystem.IsWindows() || args.Contains("--keep-alive")) return;
    while (!lifetime.ApplicationStopping.IsCancellationRequested)
    {
        await Task.Delay(TimeSpan.FromMinutes(2));
        var office = System.Diagnostics.Process.GetProcessesByName("WINWORD").Length + System.Diagnostics.Process.GetProcessesByName("EXCEL").Length;
        if (office == 0) { Log.Info("Word/Excel tidak berjalan, engine berhenti."); lifetime.StopApplication(); }
    }
});

app.Run();
