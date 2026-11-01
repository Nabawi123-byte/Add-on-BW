using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using MasterBW.Core.Util;

namespace MasterBW.Core.Ai
{
    /// <summary>Menyalakan MasterBW.Engine.exe bila belum berjalan (dipakai bersama Word &amp; Excel).</summary>
    public static class EngineLauncher
    {
        private static readonly SemaphoreSlim Gate = new SemaphoreSlim(1, 1);

        public static string FindEngineExe()
        {
            var candidates = new List<string>();
            var env = Environment.GetEnvironmentVariable("MASTERBW_ENGINE");
            if (!string.IsNullOrEmpty(env)) candidates.Add(env);
            candidates.Add(Path.Combine(AppPaths.InstallDir, "engine", "MasterBW.Engine.exe"));
            candidates.Add(Path.Combine(AppPaths.InstallDir, "MasterBW.Engine.exe"));
            var here = Path.GetDirectoryName(typeof(EngineLauncher).Assembly.Location) ?? "";
            candidates.Add(Path.Combine(here, "engine", "MasterBW.Engine.exe"));
            candidates.Add(Path.Combine(here, "MasterBW.Engine.exe"));
            return candidates.FirstOrDefault(File.Exists);
        }

        public static async Task<EngineStatus> EnsureRunningAsync(EngineClient client, Action<string> status, CancellationToken ct)
        {
            var st = await client.GetStatusAsync(ct);
            if (st.Ok) return st;
            await Gate.WaitAsync(ct);
            try
            {
                st = await client.GetStatusAsync(ct);
                if (st.Ok) return st;
                var exe = FindEngineExe();
                if (exe == null)
                    return new EngineStatus { Ok = false, Error = "MasterBW.Engine.exe tidak ditemukan. Jalankan ulang installer atau scripts\\install.ps1." };
                status?.Invoke("Menyalakan mesin AI lokal…");
                var psi = new ProcessStartInfo(exe)
                {
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WorkingDirectory = Path.GetDirectoryName(exe) ?? "",
                    WindowStyle = ProcessWindowStyle.Hidden,
                };
                Process.Start(psi);
                for (int i = 0; i < 60; i++)
                {
                    await Task.Delay(500, ct);
                    st = await client.GetStatusAsync(ct);
                    if (st.Ok) return st;
                }
                return new EngineStatus { Ok = false, Error = "Mesin AI tidak merespons. Lihat log di " + AppPaths.LogDir };
            }
            finally { Gate.Release(); }
        }
    }

}
