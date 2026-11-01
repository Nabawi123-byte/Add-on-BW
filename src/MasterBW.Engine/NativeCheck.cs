using System.Runtime.InteropServices;
using LLama.Native;
using MasterBW.Core.Util;

namespace MasterBW.Engine;

/// <summary>
/// Memastikan pustaka native llama.cpp bisa dimuat SEBELUM model dipakai,
/// lalu memberi pesan yang jelas bila gagal (penyebab tersering: Visual C++ Redistributable belum ada).
/// </summary>
public static class NativeCheck
{
    public const string VcRedistUrl = "https://aka.ms/vs/17/release/vc_redist.x64.exe";

    /// <summary>DLL runtime Visual C++ yang dibutuhkan llama.cpp (ggml memakai OpenMP → vcomp140).</summary>
    public static readonly string[] VcRuntimeDlls = { "msvcp140.dll", "vcruntime140.dll", "vcruntime140_1.dll", "vcomp140.dll" };

    public static string? Error { get; private set; }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool SetDllDirectory(string lpPathName);

    public static void Run(string baseDir)
    {
        try
        {
            if (OperatingSystem.IsWindows())
            {
                // DLL runtime yang disertakan di folder engine harus bisa ditemukan oleh llama.dll di subfolder runtimes\…
                Environment.CurrentDirectory = baseDir;
                SetDllDirectory(baseDir);
                Environment.SetEnvironmentVariable("PATH", baseDir + ";" + Environment.GetEnvironmentVariable("PATH"));
            }

            NativeLibraryConfig.All
                .WithAutoFallback(true)
                .WithLogCallback((level, message) =>
                {
                    if (level == LLamaLogLevel.Debug || level == LLamaLogLevel.Continue) return;
                    Log.Info("[llama] " + message.TrimEnd());
                });

            var missing = MissingVcRuntime(baseDir);
            if (!NativeLibraryConfig.LLama.DryRun(out var lib) || lib == null)
            {
                Error = "Mesin AI tidak bisa memuat pustaka llama.cpp (llama.dll). " +
                        (missing.Count > 0
                            ? "Penyebab: Microsoft Visual C++ Redistributable (x64) belum terpasang — tidak ditemukan " + string.Join(", ", missing) + ". "
                            : "Kemungkinan Microsoft Visual C++ Redistributable (x64) rusak/terlalu lama, atau antivirus memblokir DLL. ") +
                        "Solusi: pasang " + VcRedistUrl + " lalu tutup dan buka lagi Word/Excel. Detail ada di log: " + AppPaths.LogDir;
                Log.Error(Error);
            }
            else
            {
                Log.Info("Pustaka llama.cpp siap: " + lib.Metadata);
            }
        }
        catch (Exception ex)
        {
            Error = "Pemeriksaan pustaka AI gagal: " + ex.Message;
            Log.Error("NativeCheck", ex);
        }
    }

    public static List<string> MissingVcRuntime(string baseDir)
    {
        var missing = new List<string>();
        if (!OperatingSystem.IsWindows()) return missing;
        var sys = Environment.GetFolderPath(Environment.SpecialFolder.System);
        foreach (var dll in VcRuntimeDlls)
            if (!File.Exists(Path.Combine(baseDir, dll)) && !File.Exists(Path.Combine(sys, dll))) missing.Add(dll);
        return missing;
    }

    /// <summary>Ubah TypeInitializationException dari LLamaSharp menjadi pesan yang bisa dipahami.</summary>
    public static string Explain(Exception ex)
    {
        var root = ex;
        while (root.InnerException != null) root = root.InnerException;
        return (Error ?? "Mesin AI gagal memuat pustaka llama.cpp. Pasang Microsoft Visual C++ Redistributable x64: " + VcRedistUrl + " lalu buka ulang Word/Excel.")
               + " (Detail teknis: " + root.Message.Split('\n')[0] + ")";
    }
}
