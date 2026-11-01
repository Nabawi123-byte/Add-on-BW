using System.Diagnostics;
using System.Text;
using LLama;
using LLama.Batched;
using LLama.Common;
using LLama.Native;
using LLama.Sampling;
using MasterBW.Core.Ai;
using MasterBW.Core.Util;

namespace MasterBW.Engine;

/// <summary>
/// Memegang satu model GGUF di RAM dan melayani permintaan satu per satu.
/// Memakai cache awalan prompt (prefix cache): bagian prompt yang sama dengan permintaan sebelumnya
/// (mis. system prompt panjang berisi daftar aksi) tidak dihitung ulang — sangat mempercepat di CPU.
/// </summary>
public sealed class ModelHost : IDisposable
{
    private readonly SemaphoreSlim _gate = new(1, 1);
    private readonly MasterBWSettings _settings;
    private readonly string _modelsDir;

    private LLamaWeights? _weights;
    private BatchedExecutor? _executor;
    private Conversation? _conversation;
    private readonly List<LLamaToken> _cached = new();
    private ModelParams? _params;
    private bool _supportsThinking;
    private DateTime _lastUsed = DateTime.UtcNow;
    private Timer? _idleTimer;

    public string? ModelPath { get; private set; }
    public string? LoadError { get; private set; }
    public bool Busy => _gate.CurrentCount == 0;
    public bool Loaded => _weights != null;
    public int ContextSize => (int)(_params?.ContextSize ?? (uint)_settings.ContextSize);
    public int Threads => _params?.Threads ?? 0;

    public ModelHost(MasterBWSettings settings, string modelsDir)
    {
        _settings = settings;
        _modelsDir = modelsDir;
        if (settings.IdleUnloadMinutes > 0)
            _idleTimer = new Timer(_ => UnloadIfIdle(), null, TimeSpan.FromMinutes(1), TimeSpan.FromMinutes(1));
    }

    public string ResolveModelPath()
    {
        var file = _settings.ModelFile;
        if (Path.IsPathRooted(file) && File.Exists(file)) return file;
        var candidate = Path.Combine(_modelsDir, file);
        if (File.Exists(candidate)) return candidate;
        // fallback: GGUF pertama di folder models
        var any = Directory.Exists(_modelsDir) ? Directory.GetFiles(_modelsDir, "*.gguf").OrderBy(f => f).FirstOrDefault() : null;
        return any ?? candidate;
    }

    private void EnsureLoaded()
    {
        if (_weights != null) return;
        var path = ResolveModelPath();
        if (!File.Exists(path))
        {
            LoadError = "File model tidak ditemukan: " + path + ". Jalankan 'Unduh Model' atau salin file .gguf ke folder models.";
            throw new FileNotFoundException(LoadError);
        }
        int threads = _settings.Threads > 0 ? _settings.Threads : Math.Max(1, Environment.ProcessorCount / 2);
        _params = new ModelParams(path)
        {
            ContextSize = (uint)Math.Max(2048, _settings.ContextSize),
            GpuLayerCount = 0,
            Threads = threads,
            BatchThreads = threads,
            BatchSize = 512,
            UBatchSize = 512,
        };
        var sw = Stopwatch.StartNew();
        _weights = LLamaWeights.LoadFromFile(_params);
        _executor = new BatchedExecutor(_weights, _params);
        _conversation = null;
        _cached.Clear();
        ModelPath = path;
        LoadError = null;
        var name = (_weights.Metadata.TryGetValue("general.name", out var n) ? n : Path.GetFileName(path)) ?? "";
        var tmpl = _weights.Metadata.TryGetValue("tokenizer.chat_template", out var t) ? t : "";
        _supportsThinking = (tmpl ?? "").Contains("<think>") || name.Contains("Qwen3", StringComparison.OrdinalIgnoreCase)
                            || name.Contains("DeepSeek-R1", StringComparison.OrdinalIgnoreCase);
        Log.Info($"Model dimuat: {name} ({path}) dalam {sw.Elapsed.TotalSeconds:0.0} dtk, ctx={_params.ContextSize}, threads={threads}, thinking={_supportsThinking}");
    }

    public string ModelName => ModelPath == null ? "" : Path.GetFileNameWithoutExtension(ModelPath);

    private void UnloadIfIdle()
    {
        if (_weights == null || _settings.IdleUnloadMinutes <= 0) return;
        if (DateTime.UtcNow - _lastUsed < TimeSpan.FromMinutes(_settings.IdleUnloadMinutes)) return;
        if (!_gate.Wait(0)) return;
        try { UnloadCore(); Log.Info("Model dilepas dari RAM karena tidak dipakai."); }
        finally { _gate.Release(); }
    }

    public async Task UnloadAsync()
    {
        await _gate.WaitAsync();
        try { UnloadCore(); }
        finally { _gate.Release(); }
    }

    private void UnloadCore()
    {
        _conversation?.Dispose();
        _conversation = null;
        _executor?.Dispose();
        _executor = null;
        _weights?.Dispose();
        _weights = null;
        _cached.Clear();
    }

    // ------------------------------------------------------------------ prompt
    private string BuildPrompt(GenerateRequest req, out bool thinkingOpen)
    {
        var template = new LLamaTemplate(_weights!);
        var msgs = req.Messages.ToList();
        if (_supportsThinking && msgs.Count > 0)
        {
            // Saklar lunak Qwen3: /think atau /no_think pada pesan pengguna terakhir.
            var lastUser = msgs.FindLastIndex(m => m.Role == "user");
            if (lastUser >= 0)
                msgs[lastUser] = new ChatMessage("user", msgs[lastUser].Content + (req.Reasoning == ReasoningMode.Cepat ? " /no_think" : " /think"));
        }
        foreach (var m in msgs) template.Add(m.Role, m.Content);
        template.AddAssistant = true;
        var prompt = Encoding.UTF8.GetString(template.Apply());
        thinkingOpen = false;
        if (_supportsThinking)
        {
            if (req.Reasoning == ReasoningMode.Cepat) prompt += "<think>\n\n</think>\n\n";
            else { prompt += "<think>\n"; thinkingOpen = true; }
        }
        return prompt;
    }

    private static ISamplingPipeline MakeSampler(GenerateRequest req, bool withGrammar)
    {
        return new DefaultSamplingPipeline
        {
            Temperature = withGrammar ? Math.Min(req.Temperature, 0.3f) : req.Temperature,
            TopP = req.TopP,
            TopK = 40,
            MinP = 0.05f,
            RepeatPenalty = 1.05f,
            Grammar = withGrammar && !string.IsNullOrWhiteSpace(req.Grammar) ? new Grammar(req.Grammar, "root") : null,
        };
    }

    /// <summary>Siapkan percakapan: pakai ulang KV cache untuk awalan yang sama, lalu masukkan sisa token.</summary>
    private int PreparePrompt(LLamaToken[] tokens)
    {
        var ex = _executor!;
        if (_conversation == null || _conversation.IsDisposed || _conversation.RequiresInference || _conversation.TokenCount != _cached.Count)
        {
            _conversation?.Dispose();
            _conversation = ex.Create();
            _cached.Clear();
        }
        int common = 0;
        while (common < _cached.Count && common < tokens.Length && _cached[common] == tokens[common]) common++;
        if (common == tokens.Length) common--; // minimal satu token harus dievaluasi untuk menghasilkan logits
        if (common < 0) common = 0;
        int rewind = _cached.Count - common;
        if (rewind > 0)
        {
            _conversation.Rewind(rewind);
            _cached.RemoveRange(common, rewind);
        }
        var rest = tokens.Skip(common).ToList();
        _conversation.Prompt(rest);
        _cached.AddRange(rest);
        return common;
    }

    private void Feed(string text)
    {
        var toks = _executor!.Context.Tokenize(text, addBos: false, special: true);
        _conversation!.Prompt(toks.ToList());
        _cached.AddRange(toks);
    }

    public async Task<GenerateResponse> GenerateAsync(GenerateRequest req, Action<StreamChunk>? onChunk, CancellationToken ct)
    {
        await _gate.WaitAsync(ct);
        var sw = Stopwatch.StartNew();
        var result = new GenerateResponse();
        try
        {
            _lastUsed = DateTime.UtcNow;
            EnsureLoaded();
            var prompt = BuildPrompt(req, out var thinking);
            var tokens = _executor!.Context.Tokenize(prompt, addBos: true, special: true);
            int ctxSize = ContextSize;
            if (tokens.Length > ctxSize - 64)
                throw new InvalidOperationException($"Teks terlalu panjang untuk konteks model ({tokens.Length} token > {ctxSize}). Pilih bagian yang lebih pendek atau naikkan Context Size di Pengaturan.");

            result.PromptTokens = tokens.Length;
            result.CachedTokens = PreparePrompt(tokens);
            int budget = Math.Max(16, Math.Min(req.MaxTokens, ctxSize - tokens.Length - 16));

            var vocab = _executor.Context.NativeHandle.ModelHandle.Vocab;
            var reasoning = new StringBuilder();
            var answer = new StringBuilder();

            // ---- fase 1: berpikir (tanpa grammar) ----
            if (thinking)
            {
                int thinkBudget = req.ThinkBudget > 0 ? req.ThinkBudget : (req.Reasoning == ReasoningMode.Mendalam ? 1536 : 384);
                thinkBudget = Math.Min(thinkBudget, Math.Max(0, ctxSize - _cached.Count - budget - 32));
                using var sampler = MakeSampler(req, false);
                var decoder = new StreamingTokenDecoder(_executor.Context);
                bool closed = false;
                for (int i = 0; i < thinkBudget; i++)
                {
                    ct.ThrowIfCancellationRequested();
                    var token = await NextToken(sampler, ct);
                    if (token.IsEndOfGeneration(vocab)) break;
                    decoder.Add(token);
                    var piece = decoder.Read();
                    reasoning.Append(piece);
                    _conversation!.Prompt(token);
                    _cached.Add(token);
                    if (reasoning.ToString().Contains("</think>")) { closed = true; break; }
                    onChunk?.Invoke(new StreamChunk { Kind = "think", Delta = piece });
                }
                if (!closed)
                {
                    await EnsureInferred(ct);
                    Feed("\n</think>\n\n");
                }
                else
                {
                    await EnsureInferred(ct);
                    Feed("\n\n");
                }
                var r = reasoning.ToString();
                var idx = r.IndexOf("</think>", StringComparison.Ordinal);
                result.Reasoning = (idx >= 0 ? r[..idx] : r).Trim();
            }

            // ---- fase 2: jawaban (dengan grammar bila ada) ----
            using (var sampler = MakeSampler(req, true))
            {
                var decoder = new StreamingTokenDecoder(_executor.Context);
                int produced = 0;
                result.StopReason = "length";
                while (produced < budget)
                {
                    ct.ThrowIfCancellationRequested();
                    var token = await NextToken(sampler, ct);
                    if (token.IsEndOfGeneration(vocab)) { result.StopReason = "stop"; break; }
                    decoder.Add(token);
                    var piece = decoder.Read();
                    answer.Append(piece);
                    produced++;
                    _conversation!.Prompt(token);
                    _cached.Add(token);
                    onChunk?.Invoke(new StreamChunk { Kind = "text", Delta = piece });
                    if (req.Stop.Count > 0 && req.Stop.Any(s => answer.ToString().EndsWith(s, StringComparison.Ordinal))) { result.StopReason = "stop"; break; }
                }
                result.CompletionTokens = produced;
            }

            result.Text = ThinkFilter.StripThinking(answer.ToString()).Trim();
            foreach (var s in req.Stop) if (result.Text.EndsWith(s, StringComparison.Ordinal)) result.Text = result.Text[..^s.Length].TrimEnd();
            result.Seconds = sw.Elapsed.TotalSeconds;
            Log.Info($"Generate: prompt={result.PromptTokens} (cache {result.CachedTokens}), out={result.CompletionTokens}, {result.Seconds:0.0} dtk");
            return result;
        }
        catch (OperationCanceledException)
        {
            ResetConversation();
            throw;
        }
        catch (Exception ex)
        {
            ResetConversation();
            Log.Error("Generate gagal", ex);
            result.Error = ex.Message;
            result.Seconds = sw.Elapsed.TotalSeconds;
            return result;
        }
        finally
        {
            _lastUsed = DateTime.UtcNow;
            _gate.Release();
        }
    }

    private async Task<LLamaToken> NextToken(ISamplingPipeline sampler, CancellationToken ct)
    {
        await EnsureInferred(ct);
        return _conversation!.Sample(sampler);
    }

    private async Task EnsureInferred(CancellationToken ct)
    {
        // Infer() hanya memproses satu batch (maks. BatchSize token) per panggilan, jadi prompt panjang perlu diulang.
        while (_conversation!.RequiresInference)
        {
            ct.ThrowIfCancellationRequested();
            var res = await _executor!.Infer(ct);
            if (res != DecodeResult.Ok) throw new InvalidOperationException("Inferensi gagal: " + res + ". Coba kurangi panjang teks.");
        }
    }

    private void ResetConversation()
    {
        try { _conversation?.Dispose(); } catch { /* abaikan */ }
        _conversation = null;
        _cached.Clear();
    }

    public void Dispose()
    {
        _idleTimer?.Dispose();
        UnloadCore();
    }
}
