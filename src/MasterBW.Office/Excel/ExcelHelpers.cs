using System;
using System.Collections.Generic;
using System.Linq;
using MasterBW.Core.ExcelKit;
using MasterBW.Core.Stats;
using Excel = Microsoft.Office.Interop.Excel;

namespace MasterBW.Office.ForExcel
{
    /// <summary>Utilitas Excel: resolusi range/kolom, lembar hasil, penulisan tabel analisis.</summary>
    internal static class ExcelHelpers
    {
        public static Excel.Range Cell(Excel.Worksheet ws, int row, int col) => (Excel.Range)ws.Cells[row, col];

        public static string NumberFormat(string kind, string custom = null)
        {
            switch ((kind ?? "").ToLowerInvariant())
            {
                case "rupiah": return "\"Rp\"#,##0;[Red]-\"Rp\"#,##0";
                case "rupiah_desimal": return "\"Rp\"#,##0.00;[Red]-\"Rp\"#,##0.00";
                case "ribuan": case "angka": return "#,##0";
                case "desimal": case "desimal2": return "#,##0.00";
                case "persen": return "0.0%";
                case "tanggal": case "tanggal_pendek": return "dd/mm/yyyy";
                case "tanggal_panjang": return "[$-421]d mmmm yyyy";
                case "akuntansi": return "_(\"Rp\"* #,##0_);_(\"Rp\"* (#,##0);_(\"Rp\"* \"-\"_);_(@_)";
                case "teks": return "@";
                case "custom": return string.IsNullOrWhiteSpace(custom) ? "General" : custom;
                default: return "General";
            }
        }

        /// <summary>Ambil range dari alamat (boleh "Sheet1!A1:B5"); kosong → area data di sekitar sel aktif.</summary>
        public static Excel.Range Resolve(Excel.Application app, string address, bool fallbackToRegion = true)
        {
            var ws = (Excel.Worksheet)app.ActiveSheet;
            if (string.IsNullOrWhiteSpace(address))
            {
                if (!fallbackToRegion) return (Excel.Range)app.Selection;
                return DataRegion(app);
            }
            address = address.Trim().Replace("$", "");
            if (address.Contains("!"))
            {
                var parts = address.Split('!');
                var sheetName = parts[0].Trim('\'');
                var sheet = (Excel.Worksheet)((Excel.Workbook)app.ActiveWorkbook).Worksheets[sheetName];
                return sheet.Range[parts[1]];
            }
            // kolom saja, mis. "F" → kolom F sebatas area data
            if (System.Text.RegularExpressions.Regex.IsMatch(address, @"^[A-Za-z]{1,3}$"))
            {
                var region = DataRegion(app);
                int col = FormulaTools.ColumnNumber(address);
                return ws.Range[Cell(ws, region.Row, col), Cell(ws, region.Row + region.Rows.Count - 1, col)];
            }
            return ws.Range[address];
        }

        public static Excel.Range DataRegion(Excel.Application app)
        {
            var ws = (Excel.Worksheet)app.ActiveSheet;
            var active = (Excel.Range)app.ActiveCell;
            Excel.Range region = null;
            try
            {
                if (active != null && app.WorksheetFunction.CountA(active.CurrentRegion) > 0) region = active.CurrentRegion;
            }
            catch { }
            if (region == null || (region.Rows.Count == 1 && region.Columns.Count == 1)) region = ws.UsedRange;
            return region;
        }

        public static List<string> Headers(Excel.Range region)
        {
            var list = new List<string>();
            var first = region.Rows[1] as Excel.Range;
            var vals = first?.Value2;
            if (vals is object[,] arr)
                for (int c = arr.GetLowerBound(1); c <= arr.GetUpperBound(1); c++) list.Add(Convert.ToString(arr[arr.GetLowerBound(0), c]) ?? "");
            else list.Add(Convert.ToString(vals) ?? "");
            return list;
        }

        /// <summary>Indeks kolom (1-based, relatif terhadap region) dari huruf kolom atau nama header.</summary>
        public static int ColumnIndex(Excel.Range region, string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return 1;
            var headers = Headers(region);
            int idx = headers.FindIndex(h => string.Equals(h.Trim(), key.Trim(), StringComparison.OrdinalIgnoreCase));
            if (idx >= 0) return idx + 1;
            idx = headers.FindIndex(h => h.Trim().IndexOf(key.Trim(), StringComparison.OrdinalIgnoreCase) >= 0);
            if (idx >= 0) return idx + 1;
            if (System.Text.RegularExpressions.Regex.IsMatch(key.Trim(), @"^[A-Za-z]{1,3}$"))
            {
                int abs = FormulaTools.ColumnNumber(key);
                return Math.Max(1, abs - region.Column + 1);
            }
            if (int.TryParse(key, out var n)) return n;
            throw new ArgumentException("Kolom '" + key + "' tidak ditemukan. Kolom tersedia: " + string.Join(", ", headers));
        }

        public static object[,] Values(Excel.Range r)
        {
            var v = r.Value2;
            if (v is object[,] arr) return arr;
            var single = (object[,])Array.CreateInstance(typeof(object), new[] { 1, 1 }, new[] { 1, 1 });
            single[1, 1] = v;
            return single;
        }

        /// <summary>Kolom-kolom numerik dari range (baris pertama header bila hasHeaders).</summary>
        public static (List<string> names, List<IList<double>> cols) NumericColumns(Excel.Range r, bool hasHeaders)
        {
            var v = Values(r);
            int r0 = v.GetLowerBound(0), r1 = v.GetUpperBound(0), c0 = v.GetLowerBound(1), c1 = v.GetUpperBound(1);
            var names = new List<string>();
            var cols = new List<IList<double>>();
            for (int c = c0; c <= c1; c++)
            {
                names.Add(hasHeaders ? Convert.ToString(v[r0, c]) ?? ("Kolom " + c) : FormulaTools.ColumnLetter(r.Column + c - c0));
                var list = new List<double>();
                for (int row = hasHeaders ? r0 + 1 : r0; row <= r1; row++) list.Add(DataCleaner.ToNumber(v[row, c]));
                cols.Add(list);
            }
            return (names, cols);
        }

        public static string UniqueSheetName(Excel.Workbook wb, string baseName)
        {
            baseName = new string(baseName.Where(ch => "[]:*?/\\".IndexOf(ch) < 0).ToArray());
            if (baseName.Length > 28) baseName = baseName.Substring(0, 28);
            var existing = new HashSet<string>(wb.Worksheets.Cast<Excel.Worksheet>().Select(w => w.Name), StringComparer.OrdinalIgnoreCase);
            if (!existing.Contains(baseName)) return baseName;
            for (int i = 2; ; i++) if (!existing.Contains(baseName + " " + i)) return baseName + " " + i;
        }

        public static Excel.Worksheet AddSheet(Excel.Application app, string name)
        {
            var wb = (Excel.Workbook)app.ActiveWorkbook;
            var last = (Excel.Worksheet)wb.Worksheets[wb.Worksheets.Count];
            var ws = (Excel.Worksheet)wb.Worksheets.Add(After: last);
            ws.Name = UniqueSheetName(wb, name);
            return ws;
        }

        /// <summary>Tentukan sel kiri-atas hasil: "new_sheet" (default) atau alamat sel di sheet aktif.</summary>
        public static Excel.Range OutputAnchor(Excel.Application app, string output, string sheetName)
        {
            if (string.IsNullOrWhiteSpace(output) || output.Equals("new_sheet", StringComparison.OrdinalIgnoreCase) || output.Equals("chat", StringComparison.OrdinalIgnoreCase))
            {
                var ws = AddSheet(app, sheetName);
                return Cell(ws, 1, 1);
            }
            return Resolve(app, output, false).Cells[1, 1] as Excel.Range;
        }

        /// <summary>Tulis hasil analisis (judul, tabel, catatan, interpretasi) mulai dari anchor. Mengembalikan baris berikutnya.</summary>
        public static int WriteAnalysis(Excel.Range anchor, AnalysisResult res)
        {
            var ws = anchor.Worksheet;
            int row = anchor.Row, col = anchor.Column;
            var title = Cell(ws, row, col);
            title.Value2 = res.Title;
            title.Font.Bold = true;
            title.Font.Size = 14;
            row += 2;
            foreach (var t in res.Tables) row = WriteTable(ws, row, col, t) + 1;
            if (res.Interpretation.Count > 0)
            {
                var h = Cell(ws, row, col);
                h.Value2 = "Interpretasi";
                h.Font.Bold = true;
                row++;
                foreach (var line in res.Interpretation)
                {
                    Cell(ws, row, col).Value2 = "• " + line;
                    row++;
                }
            }
            ws.Range[Cell(ws, anchor.Row, col), Cell(ws, row, col + 10)].Columns.AutoFit();
            var first = (Excel.Range)ws.Columns[col];
            if ((double)first.ColumnWidth > 45) first.ColumnWidth = 45;
            return row;
        }

        public static int WriteTable(Excel.Worksheet ws, int row, int col, StatTable t)
        {
            var title = Cell(ws, row, col);
            title.Value2 = t.Title;
            title.Font.Bold = true;
            title.Font.Color = 0x79 * 65536 + 0x4E * 256 + 0x1F; // biru tua
            row++;
            int ncol = Math.Max(t.Headers.Count, t.Rows.Count == 0 ? 0 : t.Rows.Max(r => r.Length));
            if (ncol == 0) return row;
            var data = new object[t.Rows.Count + 1, ncol];
            for (int c = 0; c < t.Headers.Count; c++) data[0, c] = t.Headers[c];
            for (int r = 0; r < t.Rows.Count; r++)
                for (int c = 0; c < t.Rows[r].Length && c < ncol; c++)
                {
                    var v = t.Rows[r][c];
                    if (v is double d) data[r + 1, c] = double.IsNaN(d) || double.IsInfinity(d) ? (object)"-" : d;
                    else data[r + 1, c] = v;
                }
            var range = ws.Range[Cell(ws, row, col), Cell(ws, row + t.Rows.Count, col + ncol - 1)];
            range.Value2 = data;
            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders.Color = 0xBFBFBF;
            var header = ws.Range[Cell(ws, row, col), Cell(ws, row, col + ncol - 1)];
            header.Font.Bold = true;
            header.Interior.Color = 0xF7EBDD; // biru muda (BGR)
            header.WrapText = true;
            header.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            if (t.Rows.Count > 0)
            {
                var body = ws.Range[Cell(ws, row + 1, col), Cell(ws, row + t.Rows.Count, col + ncol - 1)];
                body.NumberFormat = "General";
                for (int c = 0; c < ncol; c++)
                {
                    var colRange = ws.Range[Cell(ws, row + 1, col + c), Cell(ws, row + t.Rows.Count, col + c)];
                    colRange.NumberFormat = t.ColumnFormats.TryGetValue(c, out var f) ? f : "#,##0.0000;-#,##0.0000;0";
                    if (!t.ColumnFormats.ContainsKey(c) && t.Rows.All(rw => c >= rw.Length || !(rw[c] is double dd) || Math.Abs(dd - Math.Round(dd)) < 1e-9))
                        colRange.NumberFormat = "#,##0";
                }
            }
            row += t.Rows.Count + 1;
            foreach (var n in t.Notes)
            {
                var nc = Cell(ws, row, col);
                nc.Value2 = "Catatan: " + n;
                nc.Font.Italic = true;
                nc.Font.Size = 9;
                row++;
            }
            return row;
        }
    }
}
