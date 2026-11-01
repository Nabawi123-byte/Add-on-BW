using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using MasterBW.Core.Util;

namespace MasterBW.Office.UI
{
    /// <summary>Perekam suara WAV 16 kHz mono lewat winmm (tanpa pustaka tambahan), untuk Whisper offline.</summary>
    public sealed class VoiceRecorder
    {
        [DllImport("winmm.dll", CharSet = CharSet.Unicode)]
        private static extern int mciSendString(string command, StringBuilder buffer, int bufferSize, IntPtr hwndCallback);

        private const string Alias = "masterbwrec";
        public bool IsRecording { get; private set; }

        private static void Mci(string cmd)
        {
            int err = mciSendString(cmd, null, 0, IntPtr.Zero);
            if (err != 0) throw new InvalidOperationException("Mikrofon tidak dapat dipakai (kode MCI " + err + "). Periksa izin mikrofon di Windows Settings > Privacy.");
        }

        public void Start()
        {
            try { mciSendString("close " + Alias, null, 0, IntPtr.Zero); } catch { }
            Mci("open new Type waveaudio Alias " + Alias);
            Mci("set " + Alias + " time format ms bitspersample 16 channels 1 samplespersec 16000 bytespersec 32000 alignment 2");
            Mci("record " + Alias);
            IsRecording = true;
        }

        public string Stop()
        {
            var path = Path.Combine(AppPaths.CacheDir, "suara.wav");
            try
            {
                Mci("stop " + Alias);
                if (File.Exists(path)) File.Delete(path);
                Mci("save " + Alias + " \"" + path + "\"");
            }
            finally
            {
                mciSendString("close " + Alias, null, 0, IntPtr.Zero);
                IsRecording = false;
            }
            return path;
        }
    }

    /// <summary>Hook keyboard per-thread Office untuk pintasan palet perintah (default Ctrl+Alt+Space).</summary>
    public sealed class KeyboardHook : IDisposable
    {
        private delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")] private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);
        [DllImport("user32.dll")] private static extern bool UnhookWindowsHookEx(IntPtr hhk);
        [DllImport("user32.dll")] private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")] private static extern short GetKeyState(int nVirtKey);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();

        private const int WH_KEYBOARD = 2;
        private readonly HookProc _proc;
        private IntPtr _hook = IntPtr.Zero;
        private readonly Keys _key;
        private readonly bool _ctrl, _alt, _shift;
        private readonly Action _onPressed;

        public KeyboardHook(string hotkey, Action onPressed)
        {
            _onPressed = onPressed;
            _proc = Callback;
            var parts = (string.IsNullOrWhiteSpace(hotkey) ? "Ctrl+Alt+Space" : hotkey).Split('+');
            foreach (var p in parts)
            {
                var k = p.Trim().ToLowerInvariant();
                if (k == "ctrl" || k == "control") _ctrl = true;
                else if (k == "alt") _alt = true;
                else if (k == "shift") _shift = true;
                else if (Enum.TryParse(p.Trim(), true, out Keys key)) _key = key;
            }
            if (_key == Keys.None) _key = Keys.Space;
            _hook = SetWindowsHookEx(WH_KEYBOARD, _proc, IntPtr.Zero, GetCurrentThreadId());
            if (_hook == IntPtr.Zero) Log.Warn("Hook keyboard gagal dipasang.");
        }

        private static bool Down(Keys k) => (GetKeyState((int)k) & 0x8000) != 0;

        private IntPtr Callback(int code, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                bool keyDown = ((long)lParam & 0x80000000) == 0;
                if (code >= 0 && keyDown && (Keys)(int)wParam == _key
                    && Down(Keys.ControlKey) == _ctrl && Down(Keys.Menu) == _alt && Down(Keys.ShiftKey) == _shift)
                {
                    _onPressed();
                    return (IntPtr)1;
                }
            }
            catch (Exception ex) { Log.Warn("Hook: " + ex.Message); }
            return CallNextHookEx(_hook, code, wParam, lParam);
        }

        public void Dispose()
        {
            if (_hook != IntPtr.Zero) UnhookWindowsHookEx(_hook);
            _hook = IntPtr.Zero;
        }
    }
}
