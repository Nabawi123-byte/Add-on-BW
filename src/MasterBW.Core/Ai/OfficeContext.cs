using System.Collections.Generic;
using System.Text;
using MasterBW.Core.Tools;

namespace MasterBW.Core.Ai
{
    /// <summary>Ringkasan keadaan dokumen/lembar kerja yang dikirim ke AI agar perintahnya tepat sasaran.</summary>
    public sealed class OfficeContext
    {
        public OfficeApp App { get; set; }
        public string DocumentName { get; set; }

        // Word
        public string SelectionText { get; set; }
        public int PageCount { get; set; }
        public int WordCount { get; set; }
        public List<string> Headings { get; } = new List<string>();
        public string CurrentParagraphStyle { get; set; }

        // Excel
        public string ActiveSheet { get; set; }
        public List<string> SheetNames { get; } = new List<string>();
        public string SelectionAddress { get; set; }
        public string ActiveCell { get; set; }
        public string ActiveCellFormula { get; set; }
        public string ActiveCellValue { get; set; }
        public string DataRegion { get; set; }
        /// <summary>Header kolom, mis. "A=Tanggal".</summary>
        public List<string> Headers { get; } = new List<string>();
        public List<string> SampleRows { get; } = new List<string>();
        public int DataRowCount { get; set; }

        public string ToPromptText(int maxSelectionChars = 1500)
        {
            var sb = new StringBuilder();
            if (App == OfficeApp.Word)
            {
                sb.AppendLine("Dokumen: " + (DocumentName ?? "(baru)") + ", " + PageCount + " halaman, " + WordCount + " kata.");
                if (!string.IsNullOrEmpty(CurrentParagraphStyle)) sb.AppendLine("Style paragraf aktif: " + CurrentParagraphStyle);
                if (Headings.Count > 0) sb.AppendLine("Judul-judul: " + string.Join(" | ", Headings.Count > 15 ? Headings.GetRange(0, 15) : Headings));
                if (!string.IsNullOrWhiteSpace(SelectionText))
                {
                    var sel = SelectionText.Length > maxSelectionChars ? SelectionText.Substring(0, maxSelectionChars) + " …(dipotong)" : SelectionText;
                    sb.AppendLine("Teks terpilih: \"\"\"" + sel.Trim() + "\"\"\"");
                }
                else sb.AppendLine("Tidak ada teks terpilih (kursor saja).");
            }
            else
            {
                sb.AppendLine("Workbook: " + (DocumentName ?? "(baru)") + ". Sheet aktif: " + ActiveSheet + ". Sheet lain: " + string.Join(", ", SheetNames));
                sb.AppendLine("Seleksi: " + SelectionAddress + ". Sel aktif: " + ActiveCell
                              + (string.IsNullOrEmpty(ActiveCellFormula) ? "" : " berisi rumus " + ActiveCellFormula)
                              + (string.IsNullOrEmpty(ActiveCellValue) ? "" : " nilai " + ActiveCellValue));
                if (!string.IsNullOrEmpty(DataRegion))
                {
                    sb.AppendLine("Area data: " + DataRegion + " (" + DataRowCount + " baris data, baris 1 = header).");
                    if (Headers.Count > 0) sb.AppendLine("Kolom: " + string.Join(", ", Headers));
                    if (SampleRows.Count > 0) sb.AppendLine("Contoh baris: " + string.Join(" || ", SampleRows));
                }
                else sb.AppendLine("Sheet masih kosong.");
            }
            return sb.ToString().TrimEnd();
        }
    }
}
