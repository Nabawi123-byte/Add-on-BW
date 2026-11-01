using System.Drawing;
using System.Windows.Forms;

namespace MasterBW.Office.UI
{
    internal static class Theme
    {
        public static readonly Color Primary = Color.FromArgb(31, 78, 121);
        public static readonly Color PrimaryLight = Color.FromArgb(221, 235, 247);
        public static readonly Color Back = Color.FromArgb(250, 250, 252);
        public static readonly Color UserBubble = Color.FromArgb(31, 78, 121);
        public static readonly Color Muted = Color.FromArgb(110, 110, 120);
        public static readonly Color Ok = Color.FromArgb(0, 128, 64);
        public static readonly Color Error = Color.FromArgb(192, 0, 0);
        public static readonly Font Body = new Font("Segoe UI", 9.5f);
        public static readonly Font Bold = new Font("Segoe UI", 9.5f, FontStyle.Bold);
        public static readonly Font Title = new Font("Segoe UI Semibold", 12f);
        public static readonly Font Small = new Font("Segoe UI", 8.25f);
        public static readonly Font Mono = new Font("Consolas", 9.5f);

        public static Button Button(string text, bool primary = false)
        {
            var b = new Button
            {
                Text = text,
                AutoSize = true,
                FlatStyle = FlatStyle.Flat,
                Font = Body,
                Padding = new Padding(6, 2, 6, 2),
                BackColor = primary ? Primary : Color.White,
                ForeColor = primary ? Color.White : Primary,
                Cursor = Cursors.Hand,
            };
            b.FlatAppearance.BorderColor = Primary;
            return b;
        }
    }
}
