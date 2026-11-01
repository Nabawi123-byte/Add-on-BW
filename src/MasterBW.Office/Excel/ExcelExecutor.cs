using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MasterBW.Core.Ai;
using MasterBW.Core.ExcelKit;
using MasterBW.Core.Finance;
using MasterBW.Core.Indonesia;
using MasterBW.Core.Stats;
using MasterBW.Core.Tools;
using MasterBW.Core.Util;
using MasterBW.Office.Common;
using MasterBW.Office.ForWord;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using static MasterBW.Office.ForExcel.ExcelHelpers;

namespace MasterBW.Office.ForExcel
{
    /// <summary>Menjalankan aksi dari katalog ExcelTools terhadap workbook aktif.</summary>
    public sealed class ExcelExecutor : IActionExecutor
    {
        private readonly Excel.Application _app;
        private Snapshot _snapshot;

        public ExcelExecutor(object application) { _app = (Excel.Application)application; }

        public OfficeApp App => OfficeApp.Excel;

        private Excel.Workbook Wb => _app.ActiveWorkbook == null ? _app.Workbooks.Add() : (Excel.Workbook)_app.ActiveWorkbook;
        private Excel.Worksheet Ws => (Excel.Worksheet)Wb.ActiveSheet;

        // ================================================================== konteks
        public OfficeContext GetContext()
        {
            var ctx = new OfficeContext { App = OfficeApp.Excel };
            try
            {
                var wb = Wb;
                var ws = Ws;
                ctx.DocumentName = wb.Name;
                ctx.ActiveSheet = ws.Name;
                foreach (Excel.Worksheet w in wb.Worksheets) if (w.Name != ws.Name) ctx.SheetNames.Add(w.Name);
                if (_app.Selection is Excel.Range sel) ctx.SelectionAddress = sel.Address[false, false];
                var ac = (Excel.Range)_app.ActiveCell;
                if (ac != null)
                {
                    ctx.ActiveCell = ac.Address[false, false];
                    if (ac.HasFormula is bool hf && hf) ctx.ActiveCellFormula = Convert.ToString(ac.Formula);
                    ctx.ActiveCellValue = Convert.ToString(ac.Text);
                }
                if (_app.WorksheetFunction.CountA(ws.UsedRange) > 0)
                {
                    var region = DataRegion(_app);
                    ctx.DataRegion = region.Address[false, false];
                    ctx.DataRowCount = Math.Max(0, region.Rows.Count - 1);
                    var headers = Headers(region);
                    for (int i = 0; i < headers.Count && i < 40; i++) ctx.Headers.Add(FormulaTools.ColumnLetter(region.Column + i) + "=" + headers[i]);
                    var v = Values(region);
                    for (int r = v.GetLowerBound(0) + 1; r <= Math.Min(v.GetUpperBound(0), v.GetLowerBound(0) + 2); r++)
                    {
                        var cells = new List<string>();
                        for (int c = v.GetLowerBound(1); c <= Math.Min(v.GetUpperBound(1), v.GetLowerBound(1) + 15); c++) cells.Add(Convert.ToString(v[r, c], CultureInfo.InvariantCulture));
                        ctx.SampleRows.Add(string.Join(" | ", cells));
                    }
                }
            }
            catch (Exception ex) { Log.Warn("GetContext Excel: " + ex.Message); }
            return ctx;
        }

        // ================================================================== undo (snapshot)
        private sealed class Snapshot
        {
            public string SheetName;
            public string Address;
            public object Formulas;
            public HashSet<string> Sheets;
        }

        public void BeginPlan(string title)
        {
            try
            {
                var ws = Ws;
                var used = ws.UsedRange;
                long cells = (long)used.Rows.Count * used.Columns.Count;
                _snapshot = new Snapshot
                {
                    SheetName = ws.Name,
                    Address = used.Address,
                    Formulas = cells <= 2_000_000 ? used.Formula : null,
                    Sheets = new HashSet<string>(Wb.Worksheets.Cast<Excel.Worksheet>().Select(w => w.Name)),
                };
            }
            catch (Exception ex) { Log.Warn("Snapshot gagal: " + ex.Message); _snapshot = null; }
        }

        public void EndPlan() { }

        public bool CanUndoLastPlan => _snapshot != null;

        public string UndoLastPlan()
        {
            if (_snapshot == null) return "Tidak ada aksi untuk dibatalkan.";
            var wb = Wb;
            _app.DisplayAlerts = false;
            try
            {
                foreach (var w in wb.Worksheets.Cast<Excel.Worksheet>().ToList())
                    if (!_snapshot.Sheets.Contains(w.Name)) w.Delete();
                var ws = (Excel.Worksheet)wb.Worksheets[_snapshot.SheetName];
                if (_snapshot.Formulas != null)
                {
                    ws.UsedRange.ClearContents();
                    ws.Range[_snapshot.Address].Formula = _snapshot.Formulas;
                }
                _snapshot = null;
                return "Isi sheet dikembalikan dan sheet baru dihapus. (Format sel tidak ikut dikembalikan.)";
            }
            catch (Exception ex) { return "Gagal membatalkan: " + ex.Message; }
            finally { _app.DisplayAlerts = true; }
        }

        // ================================================================== eksekusi
        public async Task<string> ExecuteAsync(ActionStep s, ExecutionSession session)
        {
            _app.ScreenUpdating = false;
            try { return await ExecuteCore(s, session); }
            finally { _app.ScreenUpdating = true; }
        }

        private async Task<string> ExecuteCore(ActionStep s, ExecutionSession session)
        {
            switch (s.Action)
            {
                // ---------------- rumus & data
                case "write_formula": return WriteFormula(s);
                case "write_values": return WriteValues(s);
                case "explain_formula":
                {
                    var cell = string.IsNullOrWhiteSpace(s.Str("cell")) ? (Excel.Range)_app.ActiveCell : Resolve(_app, s.Str("cell"), false);
                    var f = Convert.ToString(cell.Formula);
                    var text = f.StartsWith("=") ? FormulaTools.Explain(f) : "Sel " + cell.Address[false, false] + " berisi nilai, bukan rumus: " + cell.Text;
                    session.ShowText("Penjelasan rumus " + cell.Address[false, false], text);
                    return text;
                }
                case "fix_formula": return await FixFormula(s, session);
                case "terbilang":
                {
                    var src = Resolve(_app, s.Str("source"));
                    var dest = Resolve(_app, s.Str("destination"), false);
                    var v = Values(src);
                    int rows = v.GetUpperBound(0) - v.GetLowerBound(0) + 1;
                    var outArr = new object[rows, 1];
                    for (int i = 0; i < rows; i++)
                    {
                        var n = DataCleaner.ToNumber(v[v.GetLowerBound(0) + i, v.GetLowerBound(1)]);
                        outArr[i, 0] = double.IsNaN(n) ? "" : s.Bool("rupiah", true) ? Terbilang.Rupiah((decimal)n) : Terbilang.ToTitle(Terbilang.Convert(n));
                    }
                    var target = dest.Worksheet.Range[dest.Cells[1, 1], dest.Cells[rows, 1]];
                    target.Value2 = outArr;
                    return rows + " angka diubah menjadi terbilang.";
                }
                case "ai_fill": return await AiFill(s, session);

                // ---------------- format
                case "format_number":
                {
                    var r = Resolve(_app, s.Str("range"));
                    r.NumberFormat = NumberFormat(s.Str("format"), s.Str("custom"));
                    return "Format angka '" + s.Str("format") + "' diterapkan ke " + r.Address[false, false] + ".";
                }
                case "format_cells":
                {
                    var r = Resolve(_app, s.Str("range"));
                    if (s.Args["bold"] != null) r.Font.Bold = s.Bool("bold");
                    if (s.Args["italic"] != null) r.Font.Italic = s.Bool("italic");
                    if (s.Num("font_size").HasValue) r.Font.Size = s.Num("font_size").Value;
                    var fc = WordHelpers.ParseColor(s.Str("font_color")); if (fc.HasValue) r.Font.Color = fc.Value;
                    var bg = WordHelpers.ParseColor(s.Str("fill_color")); if (bg.HasValue) r.Interior.Color = bg.Value;
                    switch (s.Str("h_align")) { case "center": r.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; break; case "left": r.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft; break; case "right": r.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight; break; }
                    switch (s.Str("v_align")) { case "middle": r.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter; break; case "top": r.VerticalAlignment = Excel.XlVAlign.xlVAlignTop; break; case "bottom": r.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom; break; }
                    if (s.Args["wrap"] != null) r.WrapText = s.Bool("wrap");
                    switch (s.Str("borders"))
                    {
                        case "all": r.Borders.LineStyle = Excel.XlLineStyle.xlContinuous; break;
                        case "outline": r.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium); break;
                        case "none": r.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone; break;
                    }
                    if (s.Bool("merge")) { _app.DisplayAlerts = false; r.Merge(); _app.DisplayAlerts = true; r.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; }
                    return "Format sel diterapkan ke " + r.Address[false, false] + ".";
                }
                case "format_as_table":
                {
                    var r = Resolve(_app, s.Str("range"));
                    if (r.ListObject != null) { r.ListObject.TableStyle = s.Str("style", "TableStyleMedium2"); return "Gaya tabel diperbarui."; }
                    var lo = Ws.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, r, Type.Missing, Excel.XlYesNoGuess.xlYes);
                    lo.TableStyle = s.Str("style", "TableStyleMedium2");
                    if (!string.IsNullOrWhiteSpace(s.Str("name"))) try { lo.Name = s.Str("name").Replace(" ", "_"); } catch { }
                    r.Columns.AutoFit();
                    return "Range " + r.Address[false, false] + " dijadikan Tabel Excel.";
                }
                case "autofit":
                {
                    var r = string.IsNullOrWhiteSpace(s.Str("range")) ? Ws.UsedRange : Resolve(_app, s.Str("range"));
                    r.Columns.AutoFit();
                    r.Rows.AutoFit();
                    return "Lebar kolom disesuaikan.";
                }
                case "conditional_format": return ConditionalFormat(s);
                case "freeze_panes":
                {
                    var cell = Resolve(_app, s.Str("cell", "A2"), false);
                    Ws.Activate();
                    var win = _app.ActiveWindow;
                    win.FreezePanes = false;
                    win.ScrollRow = 1; win.ScrollColumn = 1;
                    win.SplitColumn = cell.Column - 1;
                    win.SplitRow = cell.Row - 1;
                    win.FreezePanes = true;
                    return "Panel dibekukan di " + cell.Address[false, false] + ".";
                }
                case "print_setup":
                {
                    var ps = Ws.PageSetup;
                    if (!string.IsNullOrWhiteSpace(s.Str("range"))) ps.PrintArea = Resolve(_app, s.Str("range")).Address;
                    if (s.Str("orientation") == "landscape") ps.Orientation = Excel.XlPageOrientation.xlLandscape;
                    if (s.Str("orientation") == "portrait") ps.Orientation = Excel.XlPageOrientation.xlPortrait;
                    switch ((s.Str("paper") ?? "").ToUpperInvariant())
                    {
                        case "A4": ps.PaperSize = Excel.XlPaperSize.xlPaperA4; break;
                        case "F4": ps.PaperSize = Excel.XlPaperSize.xlPaperFolio; break;
                        case "LETTER": ps.PaperSize = Excel.XlPaperSize.xlPaperLetter; break;
                        case "LEGAL": ps.PaperSize = Excel.XlPaperSize.xlPaperLegal; break;
                    }
                    if (s.Bool("fit_width", true)) { ps.Zoom = false; ps.FitToPagesWide = 1; ps.FitToPagesTall = false; }
                    if (!string.IsNullOrWhiteSpace(s.Str("title_rows"))) ps.PrintTitleRows = s.Str("title_rows");
                    ps.CenterHorizontally = true;
                    return "Pengaturan cetak diterapkan.";
                }

                // ---------------- olah data
                case "sort_range":
                {
                    var r = Resolve(_app, s.Str("range"));
                    var k1 = (Excel.Range)r.Columns[ColumnIndex(r, s.Str("key"))];
                    var o1 = s.Str("order") == "desc" ? Excel.XlSortOrder.xlDescending : Excel.XlSortOrder.xlAscending;
                    if (!string.IsNullOrWhiteSpace(s.Str("key2")))
                    {
                        var k2 = (Excel.Range)r.Columns[ColumnIndex(r, s.Str("key2"))];
                        var o2 = s.Str("order2") == "desc" ? Excel.XlSortOrder.xlDescending : Excel.XlSortOrder.xlAscending;
                        r.Sort(Key1: k1, Order1: o1, Key2: k2, Order2: o2, Header: Excel.XlYesNoGuess.xlYes);
                    }
                    else r.Sort(Key1: k1, Order1: o1, Header: Excel.XlYesNoGuess.xlYes);
                    return "Data diurutkan berdasarkan " + s.Str("key") + (o1 == Excel.XlSortOrder.xlDescending ? " (terbesar ke terkecil)." : " (terkecil ke terbesar).");
                }
                case "filter_range":
                {
                    var r = Resolve(_app, s.Str("range"));
                    int idx = ColumnIndex(r, s.Str("column"));
                    r.AutoFilter(idx, s.Str("criteria"));
                    return "Filter diterapkan pada kolom " + s.Str("column") + ": " + s.Str("criteria");
                }
                case "remove_duplicates":
                {
                    var r = Resolve(_app, s.Str("range"));
                    var cols = s.StrList("columns");
                    object colArr = cols.Count == 0
                        ? Enumerable.Range(1, r.Columns.Count).Cast<object>().ToArray()
                        : cols.Select(c => (object)ColumnIndex(r, c)).ToArray();
                    int before = (int)_app.WorksheetFunction.CountA((Excel.Range)r.Columns[1]);
                    r.RemoveDuplicates(colArr, Excel.XlYesNoGuess.xlYes);
                    int after = (int)_app.WorksheetFunction.CountA((Excel.Range)r.Columns[1]);
                    return (before - after) + " baris duplikat dihapus.";
                }
                case "clean_data":
                {
                    var r = Resolve(_app, s.Str("range"));
                    var ops = s.StrList("operations");
                    var v = Values(r);
                    int changed = DataCleaner.CleanGrid(v, ops);
                    r.Value2 = v;
                    if (ops.Contains("normalize_date") && r.Columns.Count == 1) r.NumberFormat = "dd/mm/yyyy";
                    if (ops.Contains("normalize_phone")) r.NumberFormat = "@";
                    return changed + " sel dibersihkan (" + string.Join(", ", ops) + ").";
                }
                case "split_column":
                {
                    var r = Resolve(_app, s.Str("range"));
                    var v = Values(r);
                    var rows = new List<string[]>();
                    for (int i = v.GetLowerBound(0); i <= v.GetUpperBound(0); i++) rows.Add(DataCleaner.Split(Convert.ToString(v[i, v.GetLowerBound(1)]), s.Str("delimiter")));
                    int width = Math.Max(1, rows.Max(x => x.Length));
                    var outArr = new object[rows.Count, width];
                    for (int i = 0; i < rows.Count; i++) for (int j = 0; j < rows[i].Length; j++) outArr[i, j] = rows[i][j];
                    var dest = string.IsNullOrWhiteSpace(s.Str("destination")) ? Cell(r.Worksheet, r.Row, r.Column + 1) : (Excel.Range)Resolve(_app, s.Str("destination"), false).Cells[1, 1];
                    dest.Worksheet.Range[dest, Cell(dest.Worksheet, dest.Row + rows.Count - 1, dest.Column + width - 1)].Value2 = outArr;
                    return "Kolom dipisah menjadi " + width + " kolom.";
                }
                case "combine_columns":
                {
                    var region = DataRegion(_app);
                    var ws = region.Worksheet;
                    var cols = s.StrList("columns").Select(c => ColumnIndex(region, c)).ToList();
                    var sep = s.Str("separator") ?? " ";
                    var dest = Resolve(_app, s.Str("destination"), false);
                    int destCol = dest.Column;
                    var v = Values(region);
                    int n = v.GetUpperBound(0) - v.GetLowerBound(0) + 1;
                    var outArr = new object[n, 1];
                    for (int i = 0; i < n; i++)
                    {
                        if (i == 0) { outArr[0, 0] = "Gabungan"; continue; }
                        outArr[i, 0] = string.Join(sep, cols.Select(c => Convert.ToString(v[v.GetLowerBound(0) + i, v.GetLowerBound(1) + c - 1])).Where(x => !string.IsNullOrWhiteSpace(x)));
                    }
                    ws.Range[Cell(ws, region.Row, destCol), Cell(ws, region.Row + n - 1, destCol)].Value2 = outArr;
                    return "Kolom digabung ke kolom " + FormulaTools.ColumnLetter(destCol) + ".";
                }
                case "data_validation":
                {
                    var r = Resolve(_app, s.Str("range"));
                    var val = r.Validation;
                    val.Delete();
                    var type = s.Str("type", "list");
                    if (type == "list")
                    {
                        var sepLocal = Convert.ToString(_app.International[Excel.XlApplicationInternational.xlListSeparator]);
                        val.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(sepLocal, s.StrList("values")));
                    }
                    else
                    {
                        var t = type == "whole" ? Excel.XlDVType.xlValidateWholeNumber : type == "decimal" ? Excel.XlDVType.xlValidateDecimal : type == "date" ? Excel.XlDVType.xlValidateDate : Excel.XlDVType.xlValidateTextLength;
                        val.Add(t, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, s.Str("min", "0"), s.Str("max", "999999999"));
                    }
                    if (!string.IsNullOrWhiteSpace(s.Str("message"))) { val.ErrorMessage = s.Str("message"); val.ErrorTitle = "MASTER BW"; }
                    return "Validasi data (" + type + ") dipasang di " + r.Address[false, false] + ".";
                }
                case "find_replace":
                {
                    var r = string.IsNullOrWhiteSpace(s.Str("range")) ? Ws.UsedRange : Resolve(_app, s.Str("range"));
                    r.Replace(s.Str("find"), s.Str("replace") ?? "", s.Bool("whole_cell") ? Excel.XlLookAt.xlWhole : Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, s.Bool("match_case"));
                    return "Penggantian '" + s.Str("find") + "' → '" + s.Str("replace") + "' selesai.";
                }
                case "subtotal_group": return SubtotalGroup(s);

                // ---------------- pivot, grafik, dashboard
                case "create_pivot": return CreatePivot(s);
                case "create_chart":
                {
                    var src = Resolve(_app, s.Str("source"));
                    var ws = Ws;
                    var pos = string.IsNullOrWhiteSpace(s.Str("position")) ? Cell(ws, src.Row, src.Column + src.Columns.Count + 1) : Resolve(_app, s.Str("position"), false);
                    var chart = AddChart(ws, src, s.Str("type", "column"), s.Str("title"), (double)pos.Left, (double)pos.Top, 480, 280, s.Str("x_title"), s.Str("y_title"));
                    return "Grafik " + s.Str("type") + " dibuat.";
                }
                case "create_dashboard": return CreateDashboard(s);
                case "analyze_data": return await AnalyzeData(s, session);

                // ---------------- statistik
                case "descriptive_stats":
                {
                    var (names, cols) = NumericColumns(Resolve(_app, s.Str("range")), s.Bool("has_headers", true));
                    return Output(s, Statistics.DescriptiveAnalysis(names, cols), "Deskriptif");
                }
                case "correlation":
                {
                    var (names, cols) = NumericColumns(Resolve(_app, s.Str("range")), s.Bool("has_headers", true));
                    return Output(s, Statistics.CorrelationMatrix(names, cols, s.Str("method") == "spearman"), "Korelasi");
                }
                case "regression":
                {
                    bool h = s.Bool("has_headers", true);
                    var (yn, yc) = NumericColumns(Resolve(_app, s.Str("y_range")), h);
                    var (xn, xc) = NumericColumns(Resolve(_app, s.Str("x_range")), h);
                    return Output(s, Statistics.RegressionAnalysis(yn[0], yc[0], xn, xc), "Regresi");
                }
                case "t_test":
                {
                    var r1 = Resolve(_app, s.Str("range1"));
                    bool h1 = !(r1.Cells[1, 1] is Excel.Range c1 && c1.Value2 is double);
                    var (n1, a) = NumericColumns(r1, h1);
                    List<string> n2 = null; List<IList<double>> b = null;
                    if (!string.IsNullOrWhiteSpace(s.Str("range2")))
                    {
                        var r2 = Resolve(_app, s.Str("range2"));
                        bool h2 = !(r2.Cells[1, 1] is Excel.Range c2 && c2.Value2 is double);
                        (n2, b) = NumericColumns(r2, h2);
                    }
                    else if (a.Count >= 2) { n2 = new List<string> { n1[1] }; b = new List<IList<double>> { a[1] }; }
                    var res = Statistics.TTest(s.Str("type", "independent"), a[0], b?[0], s.NumOr("mu", 0), n1[0], n2?[0] ?? "Kelompok 2");
                    return Output(s, res, "Uji t");
                }
                case "anova":
                {
                    var (names, cols) = NumericColumns(Resolve(_app, s.Str("range")), s.Bool("has_headers", true));
                    return Output(s, Statistics.AnovaAnalysis(names, cols), "ANOVA");
                }
                case "chi_square":
                {
                    var r = Resolve(_app, s.Str("range"));
                    var v = Values(r);
                    bool h = s.Bool("has_headers", true);
                    int r0 = v.GetLowerBound(0) + (h ? 1 : 0), c0 = v.GetLowerBound(1) + (h ? 1 : 0);
                    int rows = v.GetUpperBound(0) - r0 + 1, cols = v.GetUpperBound(1) - c0 + 1;
                    var obs = new double[rows, cols];
                    var rn = new List<string>(); var cn = new List<string>();
                    for (int i = 0; i < rows; i++) rn.Add(h ? Convert.ToString(v[r0 + i, v.GetLowerBound(1)]) : "Baris " + (i + 1));
                    for (int j = 0; j < cols; j++) cn.Add(h ? Convert.ToString(v[v.GetLowerBound(0), c0 + j]) : "Kolom " + (j + 1));
                    for (int i = 0; i < rows; i++) for (int j = 0; j < cols; j++) { var x = DataCleaner.ToNumber(v[r0 + i, c0 + j]); obs[i, j] = double.IsNaN(x) ? 0 : x; }
                    return Output(s, Statistics.ChiSquare(rn, cn, obs), "Chi-square");
                }
                case "normality_test":
                {
                    var (names, cols) = NumericColumns(Resolve(_app, s.Str("range")), s.Bool("has_headers", true));
                    return Output(s, Statistics.NormalityAnalysis(names, cols), "Normalitas");
                }
                case "validity_reliability":
                {
                    var (names, cols) = NumericColumns(Resolve(_app, s.Str("range")), s.Bool("has_headers", true));
                    return Output(s, Statistics.ValidityReliability(names, cols, s.NumOr("alpha", 0.05)), "Validitas Reliabilitas");
                }
                case "detect_outliers":
                {
                    var r = Resolve(_app, s.Str("range"));
                    var (names, cols) = NumericColumns(r, false);
                    var idx = Statistics.Outliers(cols[0], s.Str("method", "iqr"), out var lo, out var hi);
                    if (s.Bool("highlight", true)) foreach (var i in idx) ((Excel.Range)r.Cells[i + 1, 1]).Interior.Color = 0x9CC7FF; // oranye muda
                    var text = $"{idx.Count} outlier ditemukan (batas bawah {lo:#,##0.##}, batas atas {hi:#,##0.##}).";
                    if (idx.Count > 0) text += " Baris: " + string.Join(", ", idx.Take(30).Select(i => (r.Row + i).ToString()));
                    return text;
                }
                case "forecast": return Forecast(s);

                // ---------------- keuangan
                case "loan_schedule": return LoanSchedule(s);
                case "depreciation_schedule":
                {
                    var rows = FinanceCalc.Depreciation(s.NumOr("cost", 0), s.NumOr("salvage", 0), s.IntOr("life_years", 5), s.Str("method", "straight_line"));
                    var t = new StatTable("Jadwal Penyusutan (" + s.Str("method", "straight_line") + ")", "Tahun", "Nilai Awal", "Beban Penyusutan", "Akumulasi", "Nilai Buku");
                    foreach (var r in rows) t.Row(r.Year, r.Opening, r.Expense, r.Accumulated, r.BookValue);
                    for (int c = 1; c <= 4; c++) t.ColumnFormats[c] = NumberFormat("rupiah");
                    var res = new AnalysisResult("Penyusutan Aset Tetap"); res.Tables.Add(t);
                    res.Interpretation.Add($"Harga perolehan Rp{s.NumOr("cost", 0):#,##0}, nilai sisa Rp{s.NumOr("salvage", 0):#,##0}, umur {s.IntOr("life_years", 5)} tahun.");
                    return Output(s, res, "Penyusutan");
                }
                case "investment_analysis":
                {
                    var flows = ParseNumbers(s.Str("cashflows"));
                    double rate = s.NumOr("rate", 10);
                    var t = new StatTable("Arus Kas", "Periode", "Arus Kas", "Faktor Diskonto", "Nilai Sekarang", "Kumulatif PV");
                    double cum = 0;
                    for (int i = 0; i < flows.Count; i++) { double df = 1 / Math.Pow(1 + rate / 100, i); cum += flows[i] * df; t.Row(i, flows[i], df, flows[i] * df, cum); }
                    t.ColumnFormats[1] = t.ColumnFormats[3] = t.ColumnFormats[4] = NumberFormat("rupiah");
                    double npv = FinanceCalc.Npv(rate, flows);
                    var irr = FinanceCalc.Irr(flows);
                    var pb = FinanceCalc.PaybackPeriod(flows);
                    var dpb = FinanceCalc.PaybackPeriod(flows, rate);
                    double pi = FinanceCalc.ProfitabilityIndex(rate, flows);
                    var k = new StatTable("Kriteria Kelayakan", "Kriteria", "Nilai", "Keputusan")
                        .Row("NPV", npv, npv > 0 ? "Layak (NPV > 0)" : "Tidak layak")
                        .Row("IRR (%)", irr.HasValue ? (object)irr.Value : "-", irr.HasValue && irr > rate ? "Layak (IRR > " + rate + "%)" : "Tidak layak")
                        .Row("Profitability Index", pi, pi > 1 ? "Layak (PI > 1)" : "Tidak layak")
                        .Row("Payback Period (tahun)", pb.HasValue ? (object)pb.Value : "tidak balik modal", "")
                        .Row("Discounted Payback (tahun)", dpb.HasValue ? (object)dpb.Value : "tidak balik modal", "");
                    var res = new AnalysisResult("Analisis Kelayakan Investasi (diskonto " + rate + "%)");
                    res.Tables.Add(t); res.Tables.Add(k);
                    res.Interpretation.Add(npv > 0 ? $"Investasi LAYAK: NPV positif Rp{npv:#,##0}." : $"Investasi TIDAK LAYAK: NPV Rp{npv:#,##0}.");
                    return Output(s, res, "Investasi");
                }
                case "tax_calculation": return Tax(s);
                case "break_even":
                {
                    var (u, rp) = FinanceCalc.BreakEven(s.NumOr("fixed_cost", 0), s.NumOr("price", 0), s.NumOr("variable_cost", 0));
                    var t = new StatTable("Analisis Titik Impas (BEP)", "Komponen", "Nilai")
                        .Row("Biaya tetap", s.NumOr("fixed_cost", 0)).Row("Harga jual per unit", s.NumOr("price", 0)).Row("Biaya variabel per unit", s.NumOr("variable_cost", 0))
                        .Row("Margin kontribusi per unit", s.NumOr("price", 0) - s.NumOr("variable_cost", 0)).Row("BEP (unit)", Math.Ceiling(u)).Row("BEP (Rupiah)", rp);
                    t.ColumnFormats[1] = "#,##0";
                    var res = new AnalysisResult("Break Even Point"); res.Tables.Add(t);
                    res.Interpretation.Add($"Usaha balik modal setelah menjual {Math.Ceiling(u):#,##0} unit atau omzet Rp{rp:#,##0}.");
                    return Output(s, res, "BEP");
                }
                case "financial_ratios":
                {
                    var dict = new Dictionary<string, double>();
                    if (s.Args["values"] is JObject o) foreach (var p in o.Properties()) if (IndoNumber.TryParse(p.Value.ToString(), out var d)) dict[p.Name.ToLowerInvariant()] = d;
                    var t = new StatTable("Rasio Keuangan", "Kelompok", "Rasio", "Nilai", "Satuan", "Rumus");
                    foreach (var r in FinanceCalc.Ratios(dict)) if (r.value.HasValue) t.Row(r.group, r.name, r.value.Value, r.unit, r.formula);
                    var res = new AnalysisResult("Analisis Rasio Keuangan"); res.Tables.Add(t);
                    return Output(s, res, "Rasio");
                }

                // ---------------- templat & sheet
                case "create_template": return CreateTemplate(s);
                case "add_sheet": AddSheet(_app, s.Str("name", "Sheet Baru")).Activate(); return "Sheet ditambahkan.";
                case "rename_sheet":
                {
                    var ws = string.IsNullOrWhiteSpace(s.Str("old_name")) ? Ws : (Excel.Worksheet)Wb.Worksheets[s.Str("old_name")];
                    ws.Name = UniqueSheetName(Wb, s.Str("new_name"));
                    return "Sheet diganti nama menjadi " + ws.Name + ".";
                }
                case "delete_sheet":
                {
                    _app.DisplayAlerts = false;
                    ((Excel.Worksheet)Wb.Worksheets[s.Str("name")]).Delete();
                    _app.DisplayAlerts = true;
                    return "Sheet " + s.Str("name") + " dihapus.";
                }
                case "insert_cells":
                {
                    var at = s.Str("at"); int count = Math.Max(1, s.IntOr("count", 1));
                    if (s.Str("type") == "columns") { int c = FormulaTools.ColumnNumber(at); Ws.Range[Cell(Ws, 1, c), Cell(Ws, 1, c + count - 1)].EntireColumn.Insert(); }
                    else { int r = int.Parse(at); Ws.Range[Cell(Ws, r, 1), Cell(Ws, r + count - 1, 1)].EntireRow.Insert(); }
                    return count + " " + (s.Str("type") == "columns" ? "kolom" : "baris") + " disisipkan.";
                }
                case "delete_cells":
                {
                    if (s.Str("type") == "blank_rows")
                    {
                        var region = string.IsNullOrWhiteSpace(s.Str("at")) ? Ws.UsedRange : Resolve(_app, s.Str("at"));
                        int deleted = 0;
                        for (int i = region.Rows.Count; i >= 1; i--)
                        {
                            var row = (Excel.Range)region.Rows[i];
                            if (_app.WorksheetFunction.CountA(row) == 0) { row.EntireRow.Delete(); deleted++; }
                        }
                        return deleted + " baris kosong dihapus.";
                    }
                    var at = s.Str("at");
                    if (s.Str("type") == "columns") ((Excel.Range)Ws.Columns[at.Contains(":") ? at : at + ":" + at]).Delete();
                    else ((Excel.Range)Ws.Rows[at.Contains(":") ? at : at + ":" + at]).Delete();
                    return "Dihapus: " + at;
                }
                case "define_name":
                {
                    var r = Resolve(_app, s.Str("range"));
                    Wb.Names.Add(s.Str("name").Replace(" ", "_"), "=" + r.Address[true, true, Excel.XlReferenceStyle.xlA1, true]);
                    return "Nama range '" + s.Str("name") + "' dibuat.";
                }
                case "protect_sheet":
                {
                    if (!s.Bool("protect", true)) { Ws.Unprotect(s.Str("password") ?? ""); return "Proteksi sheet dibuka."; }
                    if (!string.IsNullOrWhiteSpace(s.Str("unlocked_range"))) Resolve(_app, s.Str("unlocked_range")).Locked = false;
                    Ws.Protect(s.Str("password") ?? "", AllowFormattingColumns: true, AllowSorting: true, AllowFiltering: true);
                    return "Sheet diproteksi" + (string.IsNullOrWhiteSpace(s.Str("password")) ? "." : " dengan kata sandi.");
                }
                case "hide_show":
                {
                    bool hide = s.Bool("hide", true);
                    var reff = s.Str("ref");
                    switch (s.Str("target"))
                    {
                        case "sheet": ((Excel.Worksheet)Wb.Worksheets[reff]).Visible = hide ? Excel.XlSheetVisibility.xlSheetHidden : Excel.XlSheetVisibility.xlSheetVisible; break;
                        case "columns": ((Excel.Range)Ws.Columns[reff]).Hidden = hide; break;
                        default: ((Excel.Range)Ws.Rows[reff]).Hidden = hide; break;
                    }
                    return (hide ? "Disembunyikan: " : "Ditampilkan: ") + reff;
                }
                case "vba_macro": return await VbaMacro(s, session);
                default: return "Aksi '" + s.Action + "' belum didukung di Excel.";
            }
        }

        // ================================================================== implementasi
        private string WriteFormula(ActionStep s)
        {
            var r = Resolve(_app, s.Str("range"), false);
            var f = FormulaTools.Normalize(s.Str("formula"));
            var syntax = FormulaTools.SyntaxProblem(f);
            if (syntax != null) throw new InvalidOperationException("Rumus tidak valid: " + syntax + " (" + f + ")");
            var compat = FormulaTools.CompatibilityIssues(f);
            if (compat.Count > 0) throw new InvalidOperationException(string.Join(" ", compat));
            if (!string.IsNullOrWhiteSpace(s.Str("header")) && r.Row > 1)
            {
                var h = Cell(r.Worksheet, r.Row - 1, r.Column);
                if (h.Value2 == null) { h.Value2 = s.Str("header"); h.Font.Bold = true; }
            }
            r.Formula = f;
            r.Columns.AutoFit();
            var first = (Excel.Range)r.Cells[1, 1];
            var txt = Convert.ToString(first.Text);
            if (txt.StartsWith("#")) return "Rumus ditulis di " + r.Address[false, false] + ", tetapi menghasilkan " + txt + ". " + FormulaTools.DiagnoseError(txt, f);
            return "Rumus " + f + " ditulis di " + r.Address[false, false] + " (hasil sel pertama: " + txt + ").";
        }

        private string WriteValues(ActionStep s)
        {
            var grid = s.Grid("values");
            if (grid.Count == 0) return "Tidak ada nilai untuk ditulis.";
            int cols = grid.Max(g => g.Count);
            var arr = new object[grid.Count, cols];
            for (int i = 0; i < grid.Count; i++)
                for (int j = 0; j < grid[i].Count; j++)
                {
                    var v = grid[i][j];
                    if (v != null && v.StartsWith("=")) arr[i, j] = FormulaTools.Normalize(v);
                    else if (v != null && Regex(v) && IndoNumber.TryParse(v, out var d)) arr[i, j] = d;
                    else arr[i, j] = v;
                }
            var start = (Excel.Range)Resolve(_app, s.Str("start_cell", "A1"), false).Cells[1, 1];
            var ws = start.Worksheet;
            var target = ws.Range[start, Cell(ws, start.Row + grid.Count - 1, start.Column + cols - 1)];
            target.Formula = arr;
            if (s.Bool("bold_header")) ((Excel.Range)target.Rows[1]).Font.Bold = true;
            target.Columns.AutoFit();
            return grid.Count + " baris × " + cols + " kolom ditulis mulai " + start.Address[false, false] + ".";
        }

        private static bool Regex(string v) => System.Text.RegularExpressions.Regex.IsMatch(v.Trim(), @"^-?(Rp\.?\s?)?[\d.,]+%?$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        private async Task<string> FixFormula(ActionStep s, ExecutionSession session)
        {
            var cell = string.IsNullOrWhiteSpace(s.Str("cell")) ? (Excel.Range)_app.ActiveCell : Resolve(_app, s.Str("cell"), false);
            var f = Convert.ToString(cell.Formula);
            if (!f.StartsWith("=")) return "Sel " + cell.Address[false, false] + " tidak berisi rumus.";
            var err = Convert.ToString(cell.Text);
            var diag = FormulaTools.DiagnoseError(err.StartsWith("#") ? err : null, f);
            session.Report("AI memeriksa rumus…");
            var ctx = GetContext().ToPromptText();
            var r = await session.Ai.GenerateAsync(Prompts.BuildFormulaFix(cell.Address[false, false], f, err, diag, ctx), session.Mode, 500, session.Ct, null, PlanGrammar.AnyObject, 0.2f);
            var o = JsonRepair.ParseObject(r.Text);
            var newF = FormulaTools.Normalize((string)o?["formula"] ?? "");
            var expl = (string)o?["explanation"] ?? diag;
            var compat = FormulaTools.CompatibilityIssues(newF);
            if (string.IsNullOrWhiteSpace(newF) || FormulaTools.SyntaxProblem(newF) != null || compat.Count > 0)
            {
                session.ShowText("Diagnosis rumus", diag + "\n\n" + expl);
                return "Diagnosis: " + diag;
            }
            var ans = MessageBox.Show("Rumus lama:\n" + f + "\n\nPenyebab: " + expl + "\n\nRumus baru:\n" + newF + "\n\nTerapkan rumus baru?", "MASTER BW – Perbaiki rumus", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (ans != DialogResult.Yes) return "Saran rumus: " + newF + " (tidak diterapkan).";
            cell.Formula = newF;
            return "Rumus diperbaiki: " + newF + " → hasil " + cell.Text;
        }

        private async Task<string> AiFill(ActionStep s, ExecutionSession session)
        {
            var src = Resolve(_app, s.Str("source"));
            var dest = (Excel.Range)Resolve(_app, s.Str("destination"), false).Cells[1, 1];
            var v = Values(src);
            var cats = s.StrList("categories");
            int n = Math.Min(v.GetUpperBound(0) - v.GetLowerBound(0) + 1, 1000);
            var cache = new Dictionary<string, string>();
            int done = 0;
            for (int i = 0; i < n; i++)
            {
                session.Ct.ThrowIfCancellationRequested();
                var val = Convert.ToString(v[v.GetLowerBound(0) + i, v.GetLowerBound(1)]);
                if (string.IsNullOrWhiteSpace(val)) continue;
                if (!cache.TryGetValue(val, out var result))
                {
                    session.Report($"AI memproses baris {i + 1}/{n}…");
                    result = await session.Ai.CompleteAsync(Prompts.BuildRowTask(s.Str("instruction"), cats, val), ReasoningMode.Cepat, 80, session.Ct);
                    result = result.Split('\n')[0].Trim();
                    cache[val] = result;
                }
                Cell(dest.Worksheet, dest.Row + i, dest.Column).Value2 = result;
                done++;
            }
            return done + " baris diisi AI di kolom " + FormulaTools.ColumnLetter(dest.Column) + ".";
        }

        private string ConditionalFormat(ActionStep s)
        {
            var r = Resolve(_app, s.Str("range"));
            int color = WordHelpers.ParseColor(s.Str("color", "merah muda")) ?? 0xCEC7FF;
            string v1 = s.Str("value1"), v2 = s.Str("value2");
            dynamic fc;
            var fcs = r.FormatConditions;
            switch (s.Str("rule"))
            {
                case "greater_than": fc = fcs.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlGreater, v1); break;
                case "less_than": fc = fcs.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlLess, v1); break;
                case "between": fc = fcs.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlBetween, v1, v2); break;
                case "equal": fc = fcs.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlEqual, IndoNumber.TryParse(v1, out _) ? v1 : "=\"" + v1 + "\""); break;
                case "text_contains": fc = fcs.Add(Excel.XlFormatConditionType.xlTextString, String: v1, TextOperator: Excel.XlContainsOperator.xlContains); break;
                case "duplicates": fc = fcs.AddUniqueValues(); fc.DupeUnique = Excel.XlDupeUnique.xlDuplicate; break;
                case "unique": fc = fcs.AddUniqueValues(); fc.DupeUnique = Excel.XlDupeUnique.xlUnique; break;
                case "top_n": fc = fcs.AddTop10(); fc.TopBottom = Excel.XlTopBottom.xlTop10Top; fc.Rank = (int)(s.Num("value1") ?? 10); break;
                case "bottom_n": fc = fcs.AddTop10(); fc.TopBottom = Excel.XlTopBottom.xlTop10Bottom; fc.Rank = (int)(s.Num("value1") ?? 10); break;
                case "above_average": fc = fcs.AddAboveAverage(); fc.AboveBelow = Excel.XlAboveBelow.xlAboveAverage; break;
                case "below_average": fc = fcs.AddAboveAverage(); fc.AboveBelow = Excel.XlAboveBelow.xlBelowAverage; break;
                case "color_scale": fcs.AddColorScale(3); return "Skala warna diterapkan.";
                case "data_bar": fcs.AddDatabar(); return "Data bar diterapkan.";
                case "icon_set": fcs.AddIconSetCondition(); return "Ikon diterapkan.";
                default: fc = fcs.Add(Excel.XlFormatConditionType.xlExpression, Formula1: FormulaTools.Normalize(v1)); break;
            }
            fc.Interior.Color = color;
            return "Format bersyarat '" + s.Str("rule") + "' diterapkan ke " + r.Address[false, false] + ".";
        }

        private string SubtotalGroup(ActionStep s)
        {
            var region = Resolve(_app, s.Str("range"));
            var ws = region.Worksheet;
            int g = ColumnIndex(region, s.Str("group_by"));
            var valueCols = s.StrList("value_columns").Select(c => ColumnIndex(region, c)).ToList();
            var fn = s.Str("function", "sum");
            var v = Values(region);
            var headers = Headers(region);
            var groups = new List<string>();
            for (int i = v.GetLowerBound(0) + 1; i <= v.GetUpperBound(0); i++)
            {
                var key = Convert.ToString(v[i, v.GetLowerBound(1) + g - 1]);
                if (!string.IsNullOrWhiteSpace(key) && !groups.Contains(key)) groups.Add(key);
            }
            groups.Sort(StringComparer.CurrentCultureIgnoreCase);
            var anchor = OutputAnchor(_app, s.Str("output", "new_sheet"), "Rekap " + headers[g - 1]);
            var outWs = anchor.Worksheet;
            int row = anchor.Row, col = anchor.Column;
            string Abs(int c) => "'" + ws.Name + "'!" + ((Excel.Range)region.Columns[c]).Offset[1, 0].Resize[region.Rows.Count - 1, 1].Address[true, true];
            var fnName = fn == "average" ? "AVERAGEIFS" : fn == "count" ? "COUNTIFS" : "SUMIFS";
            Cell(outWs, row, col).Value2 = headers[g - 1];
            for (int k = 0; k < valueCols.Count; k++) Cell(outWs, row, col + 1 + k).Value2 = (fn == "sum" ? "Total " : fn == "average" ? "Rata-rata " : fn == "count" ? "Jumlah " : fn + " ") + headers[valueCols[k] - 1];
            for (int i = 0; i < groups.Count; i++)
            {
                var keyCell = Cell(outWs, row + 1 + i, col);
                keyCell.Value2 = groups[i];
                for (int k = 0; k < valueCols.Count; k++)
                {
                    var target = Cell(outWs, row + 1 + i, col + 1 + k);
                    var keyRef = keyCell.Address[false, true];
                    if (fn == "max" || fn == "min")
                    {
                        var vals = new List<double>();
                        for (int r2 = v.GetLowerBound(0) + 1; r2 <= v.GetUpperBound(0); r2++)
                            if (Convert.ToString(v[r2, v.GetLowerBound(1) + g - 1]) == groups[i]) vals.Add(DataCleaner.ToNumber(v[r2, v.GetLowerBound(1) + valueCols[k] - 1]));
                        var clean = vals.Where(x => !double.IsNaN(x)).ToList();
                        target.Value2 = clean.Count == 0 ? 0 : fn == "max" ? clean.Max() : clean.Min();
                    }
                    else if (fn == "count") target.Formula = "=COUNTIFS(" + Abs(g) + "," + keyRef + ")";
                    else target.Formula = "=" + fnName + "(" + Abs(valueCols[k]) + "," + Abs(g) + "," + keyRef + ")";
                }
            }
            int totalRow = row + 1 + groups.Count;
            Cell(outWs, totalRow, col).Value2 = "TOTAL";
            for (int k = 0; k < valueCols.Count && fn != "average"; k++)
                Cell(outWs, totalRow, col + 1 + k).Formula = "=SUM(" + Cell(outWs, row + 1, col + 1 + k).Address[false, false] + ":" + Cell(outWs, totalRow - 1, col + 1 + k).Address[false, false] + ")";
            var table = outWs.Range[Cell(outWs, row, col), Cell(outWs, totalRow, col + valueCols.Count)];
            table.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            ((Excel.Range)table.Rows[1]).Font.Bold = true;
            ((Excel.Range)table.Rows[1]).Interior.Color = 0xF7EBDD;
            ((Excel.Range)table.Rows[table.Rows.Count]).Font.Bold = true;
            outWs.Range[Cell(outWs, row + 1, col + 1), Cell(outWs, totalRow, col + valueCols.Count)].NumberFormat = "#,##0";
            table.Columns.AutoFit();
            outWs.Activate();
            return "Rekap " + groups.Count + " kelompok dibuat dengan rumus " + fnName + " (otomatis ikut berubah bila data berubah).";
        }

        private Excel.PivotTable BuildPivot(Excel.Range src, Excel.Range dest, string name, IEnumerable<string> rows, IEnumerable<string> cols,
                                            IEnumerable<(string field, string func)> values, IEnumerable<string> filters)
        {
            var headers = Headers(src);
            string F(string key) => headers[ColumnIndex(src, key) - 1];
            var cache = Wb.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, src);
            var pt = cache.CreatePivotTable(dest, name);
            foreach (var r in rows) ((Excel.PivotField)pt.PivotFields(F(r))).Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            foreach (var c in cols) ((Excel.PivotField)pt.PivotFields(F(c))).Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
            foreach (var f in filters) ((Excel.PivotField)pt.PivotFields(F(f))).Orientation = Excel.XlPivotFieldOrientation.xlPageField;
            foreach (var (field, func) in values)
            {
                var fname = F(field);
                Excel.XlConsolidationFunction fn;
                string label;
                switch ((func ?? "sum").ToLowerInvariant())
                {
                    case "count": fn = Excel.XlConsolidationFunction.xlCount; label = "Jumlah "; break;
                    case "average": fn = Excel.XlConsolidationFunction.xlAverage; label = "Rata-rata "; break;
                    case "max": fn = Excel.XlConsolidationFunction.xlMax; label = "Maks "; break;
                    case "min": fn = Excel.XlConsolidationFunction.xlMin; label = "Min "; break;
                    default: fn = Excel.XlConsolidationFunction.xlSum; label = "Total "; break;
                }
                var df = pt.AddDataField(pt.PivotFields(fname), label + fname, fn);
                df.NumberFormat = "#,##0";
            }
            pt.RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow);
            return pt;
        }

        private string CreatePivot(ActionStep s)
        {
            var src = Resolve(_app, s.Str("source"));
            var values = new List<(string, string)>();
            if (s.Args["values"] is JArray arr)
                foreach (var v in arr)
                {
                    if (v is JObject o) values.Add(((string)o["field"], (string)o["function"] ?? "sum"));
                    else values.Add((v.ToString(), "sum"));
                }
            var anchor = OutputAnchor(_app, s.Str("destination", "new_sheet"), "Pivot");
            var pt = BuildPivot(src, anchor.Offset[2, 0], "MasterBW_Pivot_" + DateTime.Now.ToString("HHmmss"), s.StrList("rows"), s.StrList("columns"), values, s.StrList("filters"));
            anchor.Worksheet.Activate();
            return "PivotTable dibuat di sheet " + anchor.Worksheet.Name + ".";
        }

        private Excel.Chart AddChart(Excel.Worksheet ws, Excel.Range src, string type, string title, double left, double top, double width, double height, string xTitle = null, string yTitle = null)
        {
            Excel.XlChartType ct;
            switch ((type ?? "column").ToLowerInvariant())
            {
                case "bar": ct = Excel.XlChartType.xlBarClustered; break;
                case "line": ct = Excel.XlChartType.xlLineMarkers; break;
                case "pie": ct = Excel.XlChartType.xlPie; break;
                case "doughnut": ct = Excel.XlChartType.xlDoughnut; break;
                case "area": ct = Excel.XlChartType.xlArea; break;
                case "scatter": ct = Excel.XlChartType.xlXYScatter; break;
                case "column_stacked": ct = Excel.XlChartType.xlColumnStacked; break;
                default: ct = Excel.XlChartType.xlColumnClustered; break;
            }
            var shape = ws.Shapes.AddChart2(-1, ct, left, top, width, height);
            var chart = shape.Chart;
            chart.SetSourceData(src);
            if (type == "combo")
            {
                var sc = (Excel.SeriesCollection)chart.SeriesCollection();
                if (sc.Count >= 2) { var last = sc.Item(sc.Count); last.ChartType = Excel.XlChartType.xlLine; last.AxisGroup = Excel.XlAxisGroup.xlSecondary; }
            }
            if (!string.IsNullOrWhiteSpace(title)) { chart.HasTitle = true; chart.ChartTitle.Text = title; }
            try
            {
                if (!string.IsNullOrWhiteSpace(xTitle)) { var ax = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlCategory); ax.HasTitle = true; ax.AxisTitle.Text = xTitle; }
                if (!string.IsNullOrWhiteSpace(yTitle)) { var ay = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlValue); ay.HasTitle = true; ay.AxisTitle.Text = yTitle; }
            }
            catch { /* pie tidak punya sumbu */ }
            if (ct == Excel.XlChartType.xlPie || ct == Excel.XlChartType.xlDoughnut) { chart.ApplyDataLabels(Excel.XlDataLabelsType.xlDataLabelsShowPercent); }
            return chart;
        }

        private string CreateDashboard(ActionStep s)
        {
            var src = Resolve(_app, s.Str("source"));
            var srcWs = src.Worksheet;
            var headers = Headers(src);
            var v = Values(src);
            // tebak kolom ukuran (numerik terbesar) dan dimensi (teks berkardinalitas kecil) bila tidak disebut
            var numeric = new List<int>(); var textCols = new List<int>(); int dateCol = -1;
            for (int c = 1; c <= headers.Count; c++)
            {
                int nums = 0, dates = 0, n = 0; var distinct = new HashSet<string>();
                for (int r = v.GetLowerBound(0) + 1; r <= Math.Min(v.GetUpperBound(0), v.GetLowerBound(0) + 200); r++)
                {
                    var x = v[r, v.GetLowerBound(1) + c - 1];
                    if (x == null) continue;
                    n++;
                    if (x is double d) { nums++; if (d > 30000 && d < 60000 && Convert.ToString(((Excel.Range)src.Cells[2, c]).NumberFormat).Contains("y")) dates++; }
                    distinct.Add(Convert.ToString(x));
                }
                if (dates > n / 2 && dateCol < 0) dateCol = c;
                else if (nums > n / 2) numeric.Add(c);
                else if (distinct.Count <= Math.Max(30, n / 3)) textCols.Add(c);
            }
            string measure = s.Str("measure") ?? (numeric.Count > 0 ? headers[numeric.Last() - 1] : null);
            if (measure == null) return "Tidak ada kolom angka untuk dashboard.";
            int mIdx = ColumnIndex(src, measure);
            var dims = s.StrList("dimensions");
            if (dims.Count == 0) dims = textCols.Take(2).Select(c => headers[c - 1]).ToList();
            if (!string.IsNullOrWhiteSpace(s.Str("date_column"))) dateCol = ColumnIndex(src, s.Str("date_column"));

            var ws = AddSheet(_app, "Dashboard");
            ws.Activate();
            _app.ActiveWindow.DisplayGridlines = false;
            var title = Cell(ws, 1, 2);
            title.Value2 = s.Str("title", "DASHBOARD " + headers[mIdx - 1].ToUpperInvariant());
            title.Font.Size = 20; title.Font.Bold = true; title.Font.Color = 0x79 * 65536 + 0x4E * 256 + 0x1F;
            Cell(ws, 2, 2).Value2 = "Sumber: '" + srcWs.Name + "'!" + src.Address[false, false] + " • dibuat MASTER BW";
            Cell(ws, 2, 2).Font.Italic = true;

            string mRef = "'" + srcWs.Name + "'!" + ((Excel.Range)src.Columns[mIdx]).Offset[1, 0].Resize[src.Rows.Count - 1, 1].Address[true, true];
            var kpis = new[] { ("Total " + headers[mIdx - 1], "=SUM(" + mRef + ")"), ("Rata-rata", "=AVERAGE(" + mRef + ")"), ("Jumlah data", "=COUNT(" + mRef + ")"), ("Nilai tertinggi", "=MAX(" + mRef + ")") };
            for (int i = 0; i < kpis.Length; i++)
            {
                int c = 2 + i * 3;
                var box = ws.Range[Cell(ws, 4, c), Cell(ws, 6, c + 1)];
                box.Interior.Color = 0xF7EBDD;
                box.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin);
                var label = Cell(ws, 4, c); label.Value2 = kpis[i].Item1; label.Font.Bold = true; label.Font.Color = 0x595959;
                var val = ws.Range[Cell(ws, 5, c), Cell(ws, 6, c + 1)];
                val.Merge(); val.Formula = kpis[i].Item2; val.Font.Size = 18; val.Font.Bold = true; val.NumberFormat = "#,##0"; val.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            }
            ws.Range["B:M"].ColumnWidth = 12;

            int pivotCol = 16, pivotRow = 4, chartTop = 8, made = 0;
            foreach (var dim in dims.Take(3))
            {
                var pt = BuildPivot(src, Cell(ws, pivotRow, pivotCol), "Dash_" + made + "_" + DateTime.Now.ToString("HHmmss"), new[] { dim }, new string[0], new[] { (measure, "sum") }, new string[0]);
                var rng = pt.TableRange1;
                var top = (double)Cell(ws, chartTop, 2).Top;
                var left = (double)Cell(ws, chartTop, 2 + (made % 2) * 6).Left;
                if (made >= 2) top = (double)Cell(ws, chartTop + 18, 2).Top;
                var chart = AddChart(ws, rng, made == 1 ? "pie" : "column", headers[mIdx - 1] + " per " + dim, left, top, 430, 250);
                try { chart.ShowAllFieldButtons = false; } catch { }
                pivotRow += pt.TableRange1.Rows.Count + 3;
                made++;
            }
            if (dateCol > 0)
            {
                try
                {
                    var pt = BuildPivot(src, Cell(ws, pivotRow, pivotCol), "Dash_Tren_" + DateTime.Now.ToString("HHmmss"), new[] { headers[dateCol - 1] }, new string[0], new[] { (measure, "sum") }, new string[0]);
                    var dateField = (Excel.PivotField)pt.PivotFields(headers[dateCol - 1]);
                    ((Excel.Range)dateField.DataRange.Cells[1, 1]).Group(true, true, Type.Missing, new object[] { false, false, false, false, true, false, true });
                    var top = (double)Cell(ws, chartTop + 18, 2 + (made % 2) * 6).Top;
                    var chart = AddChart(ws, pt.TableRange1, "line", "Tren " + headers[mIdx - 1] + " per bulan", (double)Cell(ws, chartTop, 2 + (made % 2) * 6).Left, top, 430, 250);
                    try { chart.ShowAllFieldButtons = false; } catch { }
                    made++;
                }
                catch (Exception ex) { Log.Warn("Tren dashboard: " + ex.Message); }
            }
            Cell(ws, 3, pivotCol).Value2 = "Data pendukung (PivotTable) — klik kanan > Refresh bila data berubah";
            Cell(ws, 3, pivotCol).Font.Italic = true;
            return "Dashboard dibuat: 4 kartu KPI dan " + made + " grafik.";
        }

        private async Task<string> AnalyzeData(ActionStep s, ExecutionSession session)
        {
            var src = Resolve(_app, s.Str("range"));
            var headers = Headers(src);
            var v = Values(src);
            var sb = new StringBuilder();
            sb.AppendLine("Jumlah baris data: " + (v.GetUpperBound(0) - v.GetLowerBound(0)));
            int measure = -1;
            for (int c = 1; c <= headers.Count; c++)
            {
                var nums = new List<double>(); var texts = new Dictionary<string, int>();
                for (int r = v.GetLowerBound(0) + 1; r <= v.GetUpperBound(0); r++)
                {
                    var x = v[r, v.GetLowerBound(1) + c - 1];
                    if (x is double d) nums.Add(d); else if (x != null) { var k = Convert.ToString(x); texts[k] = texts.TryGetValue(k, out var cnt) ? cnt + 1 : 1; }
                }
                if (nums.Count > texts.Values.Sum())
                {
                    var d = Descriptive.Of(nums);
                    sb.AppendLine($"- {headers[c - 1]} (angka): total {d.Sum:#,##0.##}, rata-rata {d.Mean:#,##0.##}, median {d.Median:#,##0.##}, min {d.Min:#,##0.##}, maks {d.Max:#,##0.##}, SD {d.StdDev:#,##0.##}");
                    measure = c;
                }
                else if (texts.Count > 0)
                    sb.AppendLine($"- {headers[c - 1]} (kategori, {texts.Count} nilai unik): " + string.Join(", ", texts.OrderByDescending(t => t.Value).Take(6).Select(t => t.Key + " (" + t.Value + ")")));
            }
            if (measure > 0)
                for (int c = 1; c <= headers.Count; c++)
                {
                    if (c == measure) continue;
                    var groups = new Dictionary<string, double>();
                    for (int r = v.GetLowerBound(0) + 1; r <= v.GetUpperBound(0); r++)
                    {
                        var key = v[r, v.GetLowerBound(1) + c - 1];
                        if (key == null || key is double) continue;
                        var m = DataCleaner.ToNumber(v[r, v.GetLowerBound(1) + measure - 1]);
                        if (double.IsNaN(m)) continue;
                        var k = Convert.ToString(key);
                        groups[k] = (groups.TryGetValue(k, out var g) ? g : 0) + m;
                    }
                    if (groups.Count >= 2 && groups.Count <= 50)
                        sb.AppendLine($"- Total {headers[measure - 1]} per {headers[c - 1]}: " + string.Join(", ", groups.OrderByDescending(g => g.Value).Take(8).Select(g => g.Key + " " + g.Value.ToString("#,##0"))));
                }
            session.Report("AI menganalisis data…");
            var insight = await session.Ai.CompleteAsync(Prompts.BuildInsight(sb.ToString(), s.Str("question")), session.Mode, 900, session.Ct);
            if (s.Str("output") == "new_sheet")
            {
                var ws = AddSheet(_app, "Insight");
                Cell(ws, 1, 1).Value2 = "Analisis & Insight Data"; Cell(ws, 1, 1).Font.Bold = true; Cell(ws, 1, 1).Font.Size = 14;
                var lines = (sb + "\nInsight AI:\n" + insight).Split('\n');
                for (int i = 0; i < lines.Length; i++) Cell(ws, 3 + i, 1).Value2 = lines[i].TrimEnd('\r');
                ((Excel.Range)ws.Columns[1]).ColumnWidth = 120;
                ((Excel.Range)ws.Columns[1]).WrapText = true;
            }
            session.ShowText("Insight data", insight + "\n\n— Ringkasan statistik —\n" + sb);
            return "Analisis selesai.";
        }

        private string Forecast(ActionStep s)
        {
            var (names, cols) = NumericColumns(Resolve(_app, s.Str("y_range")), false);
            var y = cols[0].Where(d => !double.IsNaN(d)).ToList();
            int periods = Math.Max(1, s.IntOr("periods", 3));
            var fr = Statistics.ForecastSeries(y, periods, s.Str("method", "linear"), s.IntOr("window", 3));
            var t = new StatTable("Data dan Ramalan (" + fr.Equation + ")", "Periode", "Aktual", "Model/Ramalan");
            for (int i = 0; i < y.Count; i++) t.Row(i + 1, y[i], fr.Fitted[i]);
            for (int h = 0; h < periods; h++) t.Row(y.Count + h + 1, "", fr.Forecast[h]);
            t.ColumnFormats[1] = t.ColumnFormats[2] = "#,##0.##";
            var res = new AnalysisResult("Peramalan " + names[0]);
            res.Tables.Add(t);
            res.Interpretation.Add("Metode: " + fr.Equation + ". MAPE = " + (double.IsNaN(fr.Mape) ? "-" : fr.Mape.ToString("0.##") + "%") + " (makin kecil makin akurat; < 10% sangat baik).");
            res.Interpretation.Add("Ramalan " + periods + " periode ke depan: " + string.Join(", ", fr.Forecast.Select(f => f.ToString("#,##0.##"))));
            var anchor = OutputAnchor(_app, s.Str("output", "new_sheet"), "Ramalan");
            WriteAnalysis(anchor, res);
            var ws = anchor.Worksheet;
            var dataRange = ws.Range[Cell(ws, anchor.Row + 3, anchor.Column), Cell(ws, anchor.Row + 3 + y.Count + periods, anchor.Column + 2)];
            AddChart(ws, dataRange, "line", "Aktual vs Ramalan", (double)Cell(ws, anchor.Row, anchor.Column + 5).Left, (double)Cell(ws, anchor.Row + 2, 1).Top, 460, 260);
            ws.Activate();
            return "Ramalan " + periods + " periode dibuat (" + fr.Equation + ").";
        }

        private string LoanSchedule(ActionStep s)
        {
            DateTime? start = IndoDate.TryParse(s.Str("start_date"), out var d) ? d : (DateTime?)null;
            double principal = s.NumOr("principal", 0), rate = s.NumOr("annual_rate", 0);
            int months = s.IntOr("months", 12);
            var method = s.Str("method", "annuity");
            var rows = FinanceCalc.LoanSchedule(principal, rate, months, method, start);
            var t = new StatTable("Jadwal Angsuran (" + (method == "flat" ? "bunga flat" : method == "effective" ? "bunga efektif/menurun" : "anuitas") + ")",
                                  "Bulan ke-", "Tanggal", "Sisa Pokok Awal", "Angsuran", "Bunga", "Pokok", "Sisa Pokok Akhir");
            foreach (var r in rows) t.Row(r.Period, r.Date.HasValue ? IndoDate.Short(r.Date.Value) : "", r.Opening, r.Payment, r.Interest, r.Principal, r.Closing);
            t.Row("TOTAL", "", "", rows.Sum(r => r.Payment), rows.Sum(r => r.Interest), rows.Sum(r => r.Principal), "");
            for (int c = 2; c <= 6; c++) t.ColumnFormats[c] = NumberFormat("rupiah");
            var res = new AnalysisResult("Simulasi Kredit/Pinjaman");
            res.Tables.Add(new StatTable("Ringkasan", "Komponen", "Nilai").Row("Pokok pinjaman", principal).Row("Bunga per tahun (%)", rate).Row("Tenor (bulan)", months)
                .Row("Angsuran bulan pertama", rows[0].Payment).Row("Total bunga", rows.Sum(r => r.Interest)).Row("Total dibayar", rows.Sum(r => r.Payment)));
            res.Tables.Add(t);
            res.Interpretation.Add($"Angsuran per bulan {(method == "effective" ? "pertama" : "")} Rp{rows[0].Payment:#,##0}; total bunga Rp{rows.Sum(r => r.Interest):#,##0} selama {months} bulan.");
            return Output(s, res, "Angsuran");
        }

        private string Tax(ActionStep s)
        {
            List<TaxLine> lines;
            string title;
            double amount = s.NumOr("amount", 0);
            switch (s.Str("tax", "pph21"))
            {
                case "ppn": lines = TaxCalc.Ppn(amount); title = "Perhitungan PPN"; break;
                case "pph23": lines = TaxCalc.Pph23(amount, true, s.Bool("has_npwp", true)); title = "Perhitungan PPh 23"; break;
                case "pph_final_umkm": lines = TaxCalc.PphFinalUmkm(amount); title = "PPh Final UMKM 0,5%"; break;
                default: lines = TaxCalc.Pph21Annual(amount, s.Str("ptkp", "TK/0"), s.Bool("has_npwp", true)); title = "Perhitungan PPh 21 (estimasi tahunan)"; break;
            }
            var t = new StatTable(title, "Uraian", "Jumlah", "Keterangan");
            foreach (var l in lines) t.Row(l.Label, l.Amount, l.Note);
            t.ColumnFormats[1] = NumberFormat("rupiah");
            var res = new AnalysisResult(title);
            res.Tables.Add(t);
            res.Interpretation.Add("Perhitungan ini estimasi berdasarkan UU HPP; verifikasi dengan ketentuan DJP terbaru (termasuk TER PP 58/2023 untuk pemotongan bulanan).");
            return Output(s, res, "Pajak");
        }

        private string CreateTemplate(ActionStep s)
        {
            var tpl = ExcelTemplates.Find(s.Str("template"));
            if (tpl == null) return "Templat '" + s.Str("template") + "' tidak ada. Tersedia: " + string.Join(", ", ExcelTemplates.All.Select(t => t.Id));
            var ws = AddSheet(_app, tpl.Name.Length > 28 ? tpl.Id : tpl.Name);
            ws.Activate();
            _app.ActiveWindow.DisplayGridlines = false;
            var titleText = (s.Str("title") ?? tpl.Title).ToUpperInvariant();
            var settings = MasterBWSettings.Load();
            if (tpl.IsStatement) return RenderStatement(ws, tpl, titleText, settings);

            int nCols = tpl.Columns.Count;
            int rows = Math.Max(1, s.IntOr("rows", tpl.DefaultRows));
            var title = ws.Range[Cell(ws, 1, 1), Cell(ws, 1, nCols)];
            title.Merge(); title.Value2 = titleText; title.Font.Bold = true; title.Font.Size = 14; title.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            int row = 2;
            foreach (var info in tpl.InfoRows)
            {
                Cell(ws, row, 1).Value2 = info.Contains(":") && !info.EndsWith(":") ? info : info;
                ws.Range[Cell(ws, row, 1), Cell(ws, row, 2)].Merge();
                Cell(ws, row, 1).Font.Bold = true;
                if (info.StartsWith("Nama Satuan") || info.StartsWith("Unit") || info.StartsWith("Perusahaan")) Cell(ws, row, 3).Value2 = settings.Institution;
                row++;
            }
            int headerRow = row + (tpl.InfoRows.Count > 0 ? 1 : 0);
            int first = headerRow + 1, last = headerRow + rows;
            for (int c = 0; c < nCols; c++)
            {
                var col = tpl.Columns[c];
                var h = Cell(ws, headerRow, c + 1);
                h.Value2 = col.Header;
                ((Excel.Range)ws.Columns[c + 1]).ColumnWidth = col.Width;
                var letter = FormulaTools.ColumnLetter(c + 1);
                var body = ws.Range[Cell(ws, first, c + 1), Cell(ws, last, c + 1)];
                if (col.Format != "teks") body.NumberFormat = NumberFormat(col.Format);
                if (!string.IsNullOrEmpty(col.Formula))
                {
                    if (!string.IsNullOrEmpty(col.FirstRowFormula))
                    {
                        Cell(ws, first, c + 1).Formula = ExcelTemplates.Expand(col.FirstRowFormula, first, first, last, letter);
                        if (rows > 1) ws.Range[Cell(ws, first + 1, c + 1), Cell(ws, last, c + 1)].Formula = ExcelTemplates.Expand(col.Formula, first + 1, first, last, letter);
                    }
                    else body.Formula = ExcelTemplates.Expand(col.Formula, first, first, last, letter);
                }
                if (col.Choices != null && col.Choices.Length > 0)
                {
                    var sepLocal = Convert.ToString(_app.International[Excel.XlApplicationInternational.xlListSeparator]);
                    body.Validation.Delete();
                    body.Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, string.Join(sepLocal, col.Choices));
                    body.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                }
                if (col.Formula != null) body.Interior.Color = 0xF2F2F2; // kolom rumus diberi warna abu
            }
            var first0 = tpl.Columns[0].Header;
            if (tpl.AutoNumber && (first0 == "No" || first0 == "Responden"))
            {
                var nums = new object[rows, 1];
                for (int i = 0; i < rows; i++) nums[i, 0] = first0 == "No" ? (object)(i + 1) : "R-" + (i + 1);
                ws.Range[Cell(ws, first, 1), Cell(ws, last, 1)].Value2 = nums;
                ws.Range[Cell(ws, first, 1), Cell(ws, last, 1)].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            }
            int endRow = last;
            if (tpl.Columns.Any(c => c.Total != null))
            {
                endRow = last + 1;
                Cell(ws, endRow, 2).Value2 = "TOTAL";
                for (int c = 0; c < nCols; c++)
                {
                    var col = tpl.Columns[c];
                    if (col.Total == null) continue;
                    var tc = Cell(ws, endRow, c + 1);
                    tc.Formula = ExcelTemplates.Expand(col.Total, endRow, first, last, FormulaTools.ColumnLetter(c + 1));
                    tc.NumberFormat = NumberFormat(col.Format == "teks" ? "angka" : col.Format);
                }
                var totalRange = ws.Range[Cell(ws, endRow, 1), Cell(ws, endRow, nCols)];
                totalRange.Font.Bold = true;
                totalRange.Interior.Color = 0xF7EBDD;
            }
            var table = ws.Range[Cell(ws, headerRow, 1), Cell(ws, endRow, nCols)];
            table.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            table.Borders.Color = 0xA6A6A6;
            var header = ws.Range[Cell(ws, headerRow, 1), Cell(ws, headerRow, nCols)];
            header.Font.Bold = true; header.Font.Color = 0xFFFFFF; header.Interior.Color = 0x79 * 65536 + 0x4E * 256 + 0x1F;
            header.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; header.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter; header.WrapText = true;
            ((Excel.Range)ws.Rows[headerRow]).RowHeight = 30;
            int sumRow = endRow + 2;
            foreach (var (label, formula, format) in tpl.Summary)
            {
                int valueCol = nCols;
                var lc = ws.Range[Cell(ws, sumRow, Math.Max(1, nCols - 2)), Cell(ws, sumRow, nCols - 1)];
                lc.Merge(); lc.Value2 = label; lc.Font.Bold = true; lc.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                var vc = Cell(ws, sumRow, valueCol);
                vc.Formula = ExcelTemplates.Expand(formula, sumRow, first, last, FormulaTools.ColumnLetter(valueCol));
                vc.NumberFormat = NumberFormat(format);
                vc.Font.Bold = true;
                vc.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                sumRow++;
            }
            var win = _app.ActiveWindow;
            win.FreezePanes = false; win.SplitColumn = 0; win.SplitRow = headerRow; win.FreezePanes = true;
            ws.PageSetup.PrintTitleRows = "$" + headerRow + ":$" + headerRow;
            ws.PageSetup.Orientation = nCols > 8 ? Excel.XlPageOrientation.xlLandscape : Excel.XlPageOrientation.xlPortrait;
            ws.PageSetup.Zoom = false; ws.PageSetup.FitToPagesWide = 1; ws.PageSetup.FitToPagesTall = false;
            Cell(ws, first, 2).Select();
            var note = tpl.Summary.Any(x => x.formula.Contains("TERBILANG")) || tpl.Columns.Any(c => (c.Formula ?? "").Contains("TERBILANG")) ? " Fungsi TERBILANG memerlukan add-in fungsi MASTER BW (MasterBW.Functions.xll)." : "";
            return "Templat '" + tpl.Name + "' dibuat dengan " + rows + " baris isian. Kolom abu-abu berisi rumus otomatis." + note;
        }

        private string RenderStatement(Excel.Worksheet ws, SheetTemplate tpl, string titleText, MasterBWSettings settings)
        {
            ((Excel.Range)ws.Columns[1]).ColumnWidth = 4;
            ((Excel.Range)ws.Columns[2]).ColumnWidth = 44;
            ((Excel.Range)ws.Columns[3]).ColumnWidth = 22;
            var t = ws.Range["B1:C1"]; t.Merge(); t.Value2 = titleText; t.Font.Bold = true; t.Font.Size = 14; t.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            var sub = ws.Range["B2:C2"]; sub.Merge(); sub.Value2 = string.IsNullOrWhiteSpace(settings.Institution) ? "[Nama Perusahaan/Instansi]" : settings.Institution; sub.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            var per = ws.Range["B3:C3"]; per.Merge(); per.Value2 = "Periode: [isi periode]"; per.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; per.Font.Italic = true;
            int row = 5;
            var map = new Dictionary<string, string>();
            foreach (var l in tpl.Lines) { map[l.Key] = "C" + row; row++; }
            row = 5;
            foreach (var l in tpl.Lines)
            {
                var label = Cell(ws, row, 2); var val = Cell(ws, row, 3);
                label.Value2 = l.Label;
                switch (l.Kind)
                {
                    case "heading": label.Font.Bold = true; label.Interior.Color = 0xF7EBDD; val.Interior.Color = 0xF7EBDD; break;
                    case "input": label.IndentLevel = 1; val.Interior.Color = 0xCCFFFF; val.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlDot; break;
                    case "total": label.Font.Bold = true; val.Font.Bold = true; val.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous; break;
                    case "formula": label.IndentLevel = 1; break;
                }
                if (!string.IsNullOrEmpty(l.Value))
                {
                    var f = System.Text.RegularExpressions.Regex.Replace(l.Value, @"\{(\w+)\}", m => map.TryGetValue(m.Groups[1].Value, out var a) ? a : m.Value);
                    val.Formula = f;
                }
                if (l.Kind != "heading" && l.Kind != "blank") val.NumberFormat = NumberFormat(l.Format);
                row++;
            }
            Cell(ws, row + 1, 2).Value2 = "Isi sel berwarna kuning; sel lain dihitung otomatis.";
            Cell(ws, row + 1, 2).Font.Italic = true;
            Cell(ws, row + 1, 2).Font.Size = 9;
            ws.Range["C5"].Select();
            return "Templat '" + tpl.Name + "' dibuat. Isi sel kuning, total dihitung otomatis.";
        }

        private async Task<string> VbaMacro(ActionStep s, ExecutionSession session)
        {
            var code = VbaSafety.ExtractCode(s.Str("code"));
            if (string.IsNullOrWhiteSpace(code) || !code.Contains("Sub "))
            {
                session.Report("AI menulis kode VBA…");
                var r = await session.Ai.CompleteAsync(Prompts.BuildVba(s.Str("description"), OfficeApp.Excel, GetContext().ToPromptText()), session.Mode, 1500, session.Ct);
                code = VbaSafety.ExtractCode(r);
            }
            var risks = VbaSafety.Scan(code);
            var msg = "Makro: " + s.Str("description") + "\n\n" + (risks.Count > 0 ? "⚠ PERINGATAN — kode ini: " + string.Join("; ", risks) + "\n\n" : "Pemeriksaan keamanan: tidak ditemukan perintah berbahaya.\n\n") +
                      "Pasang makro ke workbook ini?";
            session.ShowText("Kode VBA", code);
            if (MessageBox.Show(msg, "MASTER BW – Makro VBA", MessageBoxButtons.YesNo, risks.Count > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Question) != DialogResult.Yes)
                return "Kode VBA ditampilkan tetapi tidak dipasang.";
            try
            {
                dynamic vbproj = ((dynamic)Wb).VBProject;
                dynamic module = vbproj.VBComponents.Add(1);
                module.Name = "MasterBW_" + DateTime.Now.ToString("HHmmss");
                module.CodeModule.AddFromString(code);
                var sub = VbaSafety.FirstSubName(code);
                if (s.Bool("run_now") && sub != null) _app.Run(sub);
                return "Makro dipasang di modul " + (string)module.Name + (sub != null ? " (jalankan: Alt+F8 → " + sub + ")" : "") + ". Simpan workbook sebagai .xlsm agar makro tersimpan.";
            }
            catch (Exception ex)
            {
                Log.Warn("Pasang VBA gagal: " + ex.Message);
                Clipboard.SetText(code);
                return "Tidak bisa memasang otomatis (aktifkan File > Options > Trust Center > Trust Center Settings > Macro Settings > 'Trust access to the VBA project object model'). Kode sudah disalin ke clipboard: tekan Alt+F11 > Insert > Module > Ctrl+V.";
            }
        }

        // ---------------------------------------------------------------- pembantu
        private string Output(ActionStep s, AnalysisResult res, string sheetName)
        {
            var anchor = OutputAnchor(_app, s.Str("output", "new_sheet"), sheetName);
            WriteAnalysis(anchor, res);
            anchor.Worksheet.Activate();
            return res.Title + " selesai (sheet " + anchor.Worksheet.Name + "). " + string.Join(" ", res.Interpretation.Take(2));
        }

        private List<double> ParseNumbers(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return new List<double>();
            if (System.Text.RegularExpressions.Regex.IsMatch(text.Trim(), @"^('?[^!]+'?!)?\$?[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?$"))
            {
                var v = Values(Resolve(_app, text));
                var list = new List<double>();
                foreach (var x in v) { var d = DataCleaner.ToNumber(x); if (!double.IsNaN(d)) list.Add(d); }
                return list;
            }
            return text.Split(new[] { ';', '|', '\n' }.Concat(text.Contains(";") ? new char[0] : new[] { ',' }).ToArray(), StringSplitOptions.RemoveEmptyEntries)
                       .Select(t => IndoNumber.TryParse(t.Trim(), out var d) ? d : double.NaN).Where(d => !double.IsNaN(d)).ToList();
        }
    }
}
