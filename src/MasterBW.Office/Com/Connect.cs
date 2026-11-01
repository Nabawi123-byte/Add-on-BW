using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using MasterBW.Core.Tools;
using MasterBW.Core.Util;
using MasterBW.Office.UI;

[assembly: ComVisible(false)]

namespace MasterBW.Office.Com
{
    /// <summary>
    /// Wadah ActiveX untuk panel samping (Custom Task Pane). Office membuatnya lewat ProgID "MasterBW.TaskPane".
    /// </summary>
    [ComVisible(true)]
    [Guid("a0ec4d74-04ef-4d41-9403-139c35f56883")]
    [ProgId("MasterBW.TaskPane")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class TaskPaneHost : UserControl
    {
        public TaskPaneHost() { BackColor = Theme.Back; }

        public void Attach(Control content)
        {
            Controls.Clear();
            content.Dock = DockStyle.Fill;
            Controls.Add(content);
        }

        public Control Content => Controls.Count > 0 ? Controls[0] : null;
    }

    /// <summary>Dasar COM add-in Office: koneksi, ribbon, dan panel samping per jendela.</summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public abstract class ConnectBase : IDTExtensibility2, IRibbonExtensibility, ICustomTaskPaneConsumer
    {
        static ConnectBase()
        {
            // Office memuat DLL ini dari folder instalasi; pastikan dependensi (Core, Newtonsoft) dicari di folder yang sama.
            AppDomain.CurrentDomain.AssemblyResolve += ResolveFromOwnFolder;
        }

        private static Assembly ResolveFromOwnFolder(object sender, ResolveEventArgs args)
        {
            try
            {
                var dir = Path.GetDirectoryName(typeof(ConnectBase).Assembly.Location);
                var name = new AssemblyName(args.Name).Name;
                var path = Path.Combine(dir ?? "", name + ".dll");
                return File.Exists(path) ? Assembly.LoadFrom(path) : null;
            }
            catch { return null; }
        }

        protected abstract OfficeApp App { get; }

        private object _application;
        private object _ctpFactory;
        private object _ribbon;
        private object _controller; // AddInController (disimpan sebagai object agar static ctor jalan duluan)
        private readonly Dictionary<string, object> _panes = new Dictionary<string, object>();

        private AddInController Controller => (AddInController)_controller;

        // ------------------------------------------------------------------ IDTExtensibility2
        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                _application = Application;
                ((dynamic)AddInInst).Object = this;
                Init();
            }
            catch (Exception ex) { Log.Error("OnConnection", ex); }
        }

        private void Init()
        {
            var ctl = new AddInController(App, _application);
            ctl.ShowPane = ShowPane;
            ctl.TogglePane = TogglePane;
            _controller = ctl;
            ctl.Startup();
            if (App == OfficeApp.Excel) RegisterFunctionsXll();
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try { Controller?.Dispose(); } catch { }
            foreach (var p in _panes.Values) { try { ((dynamic)p).Delete(); } catch { } }
            _panes.Clear();
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        // ------------------------------------------------------------------ Ribbon
        public string GetCustomUI(string RibbonID)
        {
            try
            {
                if (App == OfficeApp.Word && RibbonID != "Microsoft.Word.Document") return null;
                if (_controller == null) Init();
                return Controller.RibbonXml();
            }
            catch (Exception ex) { Log.Error("GetCustomUI", ex); return null; }
        }

        public void OnLoad(object ribbon) { _ribbon = ribbon; }

        public void OnAction(object control)
        {
            try
            {
                dynamic c = control;
                Controller.OnCommand((string)c.Tag);
            }
            catch (Exception ex) { Log.Error("OnAction", ex); MessageBox.Show(ex.Message, "MASTER BW"); }
        }

        // ------------------------------------------------------------------ Task pane
        public void CTPFactoryAvailable(object CTPFactoryInst) { _ctpFactory = CTPFactoryInst; }

        private string WindowKey(out object window)
        {
            dynamic app = _application;
            window = app.ActiveWindow;
            if (App == OfficeApp.Excel) return "xl" + (int)((dynamic)window).Hwnd;
            dynamic w = window;
            return "wd" + (string)w.Caption + "|" + (int)w.WindowNumber;
        }

        private ChatPane ShowPane()
        {
            try
            {
                var key = WindowKey(out var window);
                if (_panes.TryGetValue(key, out var existing))
                {
                    try
                    {
                        dynamic p = existing;
                        p.Visible = true;
                        return (ChatPane)((TaskPaneHost)p.ContentControl).Content;
                    }
                    catch { _panes.Remove(key); }
                }
                if (_ctpFactory == null) throw new InvalidOperationException("Office belum menyediakan panel samping.");
                dynamic ctp = ((dynamic)_ctpFactory).CreateCTP("MasterBW.TaskPane", "MASTER BW", window);
                var host = (TaskPaneHost)ctp.ContentControl;
                var chat = Controller.CreateChatPane();
                host.Attach(chat);
                ctp.Width = 400;
                ctp.Visible = true;
                _panes[key] = ctp;
                return chat;
            }
            catch (Exception ex)
            {
                Log.Error("ShowPane", ex);
                MessageBox.Show("Panel MASTER BW gagal dibuka: " + ex.Message + "\n\nJalankan ulang install.ps1 agar komponen panel terdaftar.", "MASTER BW");
                return null;
            }
        }

        private void TogglePane()
        {
            var key = WindowKey(out _);
            if (_panes.TryGetValue(key, out var existing))
            {
                try { dynamic p = existing; p.Visible = !(bool)p.Visible; return; }
                catch { _panes.Remove(key); }
            }
            ShowPane();
        }

        private void RegisterFunctionsXll()
        {
            try
            {
                var dir = Path.GetDirectoryName(typeof(ConnectBase).Assembly.Location) ?? "";
                dynamic app = _application;
                bool is64 = Environment.Is64BitProcess;
                var xll = Path.Combine(dir, "functions", is64 ? "MasterBW.Functions-AddIn64-packed.xll" : "MasterBW.Functions-AddIn-packed.xll");
                if (!File.Exists(xll)) xll = Path.Combine(dir, "functions", is64 ? "MasterBW.Functions-AddIn64.xll" : "MasterBW.Functions-AddIn.xll");
                if (File.Exists(xll)) app.RegisterXLL(xll);
                else Log.Warn("Add-in fungsi sel tidak ditemukan: " + xll);
            }
            catch (Exception ex) { Log.Warn("RegisterXLL: " + ex.Message); }
        }
    }

    [ComVisible(true)]
    [Guid("5d44520c-7476-4ff9-add1-bfedace586de")]
    [ProgId("MasterBW.WordAddIn")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class WordConnect : ConnectBase
    {
        protected override OfficeApp App => OfficeApp.Word;
    }

    [ComVisible(true)]
    [Guid("0e03ea87-46e9-40aa-aeea-16c471df42a7")]
    [ProgId("MasterBW.ExcelAddIn")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class ExcelConnect : ConnectBase
    {
        protected override OfficeApp App => OfficeApp.Excel;
    }
}
