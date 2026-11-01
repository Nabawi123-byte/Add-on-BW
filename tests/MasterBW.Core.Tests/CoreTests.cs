using MasterBW.Core.Academic;
using MasterBW.Core.Ai;
using MasterBW.Core.ExcelKit;
using MasterBW.Core.Finance;
using MasterBW.Core.Indonesia;
using MasterBW.Core.Stats;
using MasterBW.Core.Tools;
using Xunit;

namespace MasterBW.Core.Tests;

public class IndonesiaTests
{
    [Theory]
    [InlineData(0, "nol")]
    [InlineData(11, "sebelas")]
    [InlineData(19, "sembilan belas")]
    [InlineData(100, "seratus")]
    [InlineData(1000, "seribu")]
    [InlineData(1250000, "satu juta dua ratus lima puluh ribu")]
    [InlineData(2_000_015, "dua juta lima belas")]
    [InlineData(1_000_000_000, "satu miliar")]
    public void Terbilang_Benar(long n, string expected) => Assert.Equal(expected, Terbilang.Convert((decimal)n));

    [Fact]
    public void Terbilang_Rupiah_DenganSen() =>
        Assert.Equal("Satu Juta Lima Ratus Ribu Rupiah Lima Puluh Sen", Terbilang.Rupiah(1_500_000.50m));

    [Theory]
    [InlineData("Rp 1.250.000,50", 1250000.50)]
    [InlineData("Rp1.250.000", 1250000)]
    [InlineData("1,250,000.75", 1250000.75)]
    [InlineData("1,5 jt", 1500000)]
    [InlineData("2 M", 2_000_000_000)]
    [InlineData("15%", 0.15)]
    [InlineData("(1.000)", -1000)]
    [InlineData("Rp 50.000,-", 50000)]
    public void IndoNumber_Parse(string input, double expected)
    {
        Assert.True(IndoNumber.TryParse(input, out var v));
        Assert.Equal(expected, v, 6);
    }

    [Theory]
    [InlineData("12 Januari 2024", 2024, 1, 12)]
    [InlineData("Senin, 5 Feb 2024", 2024, 2, 5)]
    [InlineData("12-01-2024", 2024, 1, 12)]
    [InlineData("3/4/24", 2024, 4, 3)]
    [InlineData("2024-08-17", 2024, 8, 17)]
    [InlineData("17 Agustus 1945", 1945, 8, 17)]
    public void IndoDate_Parse(string input, int y, int m, int d)
    {
        Assert.True(IndoDate.TryParse(input, out var date));
        Assert.Equal(new DateTime(y, m, d), date);
    }

    [Fact]
    public void Roman_BolakBalik()
    {
        Assert.Equal("MMXXVI", Roman.ToRoman(2026));
        Assert.Equal(1994, Roman.FromRoman("MCMXCIV"));
    }

    [Theory]
    [InlineData("0812-3456-7890", "+6281234567890")]
    [InlineData("+62 812 3456 7890", "+6281234567890")]
    public void NormalisasiHp(string input, string expected) => Assert.Equal(expected, IndoText.NormalizePhone(input));

    [Fact]
    public void NamaRapi_GelarDipertahankan() =>
        Assert.Equal("Siti Aminah S.Pd.", IndoText.ProperName("SITI AMINAH s.pd."));

    [Fact]
    public void Eyd_DeteksiKataTidakBaku_DanKataDepan()
    {
        var text = "kami melakukan analisa resiko di sekolah. Data di tulis diatas kertas , lalu disimpan.";
        var issues = EydChecker.Check(text);
        Assert.Contains(issues, i => i.Original == "analisa" && i.Suggestion == "analisis");
        Assert.Contains(issues, i => i.Original == "resiko" && i.Suggestion == "risiko");
        Assert.Contains(issues, i => i.Original == "di tulis" && i.Suggestion == "ditulis");
        Assert.Contains(issues, i => i.Original == "diatas" && i.Suggestion == "di atas");
        Assert.Contains(issues, i => i.Category == "spasi");
        Assert.Contains(issues, i => i.Category == "kapital" && i.Original == "k");
        Assert.DoesNotContain(issues, i => i.Original == "di sekolah");
        var fixedText = EydChecker.ApplyAll(text, issues);
        Assert.Contains("analisis risiko", fixedText);
        Assert.Contains("ditulis di atas kertas,", fixedText);
    }

    [Fact]
    public void Eyd_Rupiah_Dan_Persen()
    {
        var issues = EydChecker.Check("Harga Rp. 10.000 naik 50 %.");
        Assert.Contains(issues, i => i.Suggestion == "Rp");
        Assert.Contains(issues, i => i.Suggestion == "%");
    }
}

public class StatsTests
{
    private static readonly double[] X = { 12, 15, 11, 18, 20, 22, 14, 16, 19, 25, 13, 17, 21, 23, 16 };
    private static readonly double[] Y = { 2.1, 2.9, 2.2, 3.8, 4.1, 4.5, 2.6, 3.3, 3.7, 5.0, 2.4, 3.4, 4.2, 4.6, 3.1 };

    [Fact]
    public void Deskriptif_SamaDenganExcel()
    {
        var d = Descriptive.Of(X);
        Assert.Equal(17.4666666667, d.Mean, 6);
        Assert.Equal(0.19089376226586233, d.Skewness, 6);
        Assert.Equal(-0.8637340853918376, d.Kurtosis, 6);
        Assert.Equal(17, d.Median);
        Assert.Equal(16, d.Mode);
    }

    [Fact]
    public void Korelasi_Pearson_Spearman()
    {
        Assert.Equal(0.9921085972656644, Statistics.Pearson(X, Y), 9);
        Assert.Equal(4.136576657005164e-13, Statistics.PearsonP(Statistics.Pearson(X, Y), X.Length), 14);
        Assert.Equal(0.9919575006559455, Statistics.Spearman(X, Y), 9);
    }

    [Fact]
    public void Distribusi_TKritis() => Assert.Equal(2.0484071417952454, Distributions.TInvTwoTailed(0.05, 28), 6);

    [Fact]
    public void UjiT_Independen_Berpasangan()
    {
        var ind = Statistics.TTest("independent", X, Y);
        Assert.Contains(ind.Tables[1].Rows[0], v => v is double d && Math.Abs(d - 12.70115053947338) < 1e-6);
        var pair = Statistics.TTest("paired", X, Y);
        Assert.Contains(pair.Tables[1].Rows[0], v => v is double d && Math.Abs(d - 16.593426280498196) < 1e-6);
        Assert.Contains(pair.Tables[1].Rows[0], v => v is double d && Math.Abs(d - 1.3290586796213664e-10) < 1e-12);
    }

    [Fact]
    public void Anova_SamaDenganScipy()
    {
        var a = Statistics.OneWay(new List<IList<double>> { new double[] { 1, 2, 3, 4, 5 }, new double[] { 2, 3, 4, 5, 9 }, new double[] { 5, 6, 7, 8, 9 } });
        Assert.Equal(4.943089430894308, a.F, 9);
        Assert.Equal(0.027168634522243884, a.P, 9);
    }

    [Fact]
    public void ChiSquare_SamaDenganScipy()
    {
        var r = Statistics.ChiSquare(new[] { "A", "B" }, new[] { "x", "y", "z" }, new double[,] { { 10, 20, 30 }, { 20, 15, 5 } });
        var row = r.Tables[1].Rows[0];
        Assert.Equal(18.650793650793652, (double)row[0], 9);
        Assert.Equal(8.913158246593224e-05, (double)row[2], 10);
    }

    [Fact]
    public void Regresi_Berganda_SamaDenganStatsmodels()
    {
        double[] x2 = { 1, 3, 2, 5, 4, 6, 2, 3, 5, 7, 1, 4, 6, 6, 3 };
        var m = Statistics.Ols(Y, new List<IList<double>> { X, x2 });
        Assert.Equal(0.00131698, m.Coef[0], 6);
        Assert.Equal(0.17872389, m.Coef[1], 6);
        Assert.Equal(0.08714805, m.Coef[2], 6);
        Assert.Equal(0.9874190150964026, m.R2, 9);
        Assert.Equal(470.91019788795705, m.F, 5);
        Assert.Equal(5.52589935e-06, m.P[1], 9);
        Assert.Equal(0.109152332, m.P[2], 6);
    }

    [Theory]
    [InlineData(new double[] { 12, 15, 11, 18, 20, 22, 14, 16, 19, 25, 13, 17, 21, 23, 16 }, 0.9773761296894404, 0.9484209407711799)]
    [InlineData(new double[] { 1, 2, 4, 7, 20 }, 0.8015315347624348, 0.08339219148094756)]
    [InlineData(new double[] { 1, 1, 1, 2, 2, 3, 10, 12, 15, 30, 1, 2, 1, 1, 50, 2, 3, 1, 1, 2 }, 0.5603098535517894, 1.1735266413106346e-06)]
    public void ShapiroWilk_MendekatiScipy(double[] data, double w, double p)
    {
        var (W, P) = Statistics.ShapiroWilk(data.OrderBy(v => v).ToArray());
        Assert.Equal(w, W, 3);
        Assert.True(Math.Abs(P - p) < Math.Max(0.01, p * 0.1), $"p={P} vs {p}");
    }

    [Fact]
    public void Lilliefors_MendekatiStatsmodels()
    {
        var r = Statistics.Normality(X);
        Assert.Equal(0.10406386917148763, r.KsD, 6);
        Assert.True(r.KsP > 0.2);
    }

    [Fact]
    public void Reliabilitas_CronbachAlpha()
    {
        // 5 responden × 3 item yang konsisten → alpha tinggi
        var items = new List<IList<double>> { new double[] { 4, 5, 3, 2, 5 }, new double[] { 4, 4, 3, 2, 5 }, new double[] { 5, 5, 2, 1, 4 } };
        var res = Statistics.ValidityReliability(new[] { "P1", "P2", "P3" }, items);
        var alpha = (double)res.Tables[1].Rows[0][0];
        Assert.InRange(alpha, 0.9, 1.0);
        Assert.Equal(0.8783, Distributions.RTable(5), 3);
    }

    [Fact]
    public void Peramalan_Linear()
    {
        var f = Statistics.ForecastSeries(new double[] { 10, 20, 30, 40 }, 2);
        Assert.Equal(50, f.Forecast[0], 6);
        Assert.Equal(60, f.Forecast[1], 6);
    }
}

public class FinanceTests
{
    [Fact]
    public void Anuitas_SamaDenganPMTExcel()
    {
        // PMT(1%,12,1000000) = 88.848,79
        Assert.Equal(88848.79, FinanceCalc.Pmt(0.01, 12, 1_000_000), 2);
        var rows = FinanceCalc.LoanSchedule(1_000_000, 12, 12);
        Assert.Equal(0, rows.Last().Closing, 6);
        Assert.Equal(1_000_000, rows.Sum(r => r.Principal), 4);
    }

    [Fact]
    public void Npv_Irr()
    {
        var cf = new double[] { -1000, 300, 400, 500 };
        Assert.Equal(-21.036814425244188, FinanceCalc.Npv(10, cf), 6);
        Assert.Equal(8.896, FinanceCalc.Irr(cf)!.Value, 2);
        Assert.Equal(2.6, FinanceCalc.PaybackPeriod(cf)!.Value, 6);
    }

    [Fact]
    public void Penyusutan_GarisLurus()
    {
        var rows = FinanceCalc.Depreciation(50_000_000, 5_000_000, 5);
        Assert.All(rows, r => Assert.Equal(9_000_000, r.Expense, 4));
        Assert.Equal(5_000_000, rows.Last().BookValue, 4);
    }

    [Fact]
    public void Pajak_Pasal17_Dan_Ptkp()
    {
        Assert.Equal(63_000_000, TaxCalc.Ptkp("K/1"));
        Assert.Equal(3_000_000 + 28_500_000, TaxCalc.Pasal17(250_000_000), 2);
        var lines = TaxCalc.Pph21Annual(10_000_000, "TK/0");
        // bruto 120 jt, biaya jabatan 6 jt, neto 114 jt, PTKP 54 jt, PKP 60 jt → 5% = 3 jt
        Assert.Equal(3_000_000, lines.First(l => l.Label.StartsWith("Total PPh 21")).Amount, 2);
    }

    [Fact]
    public void Ppn_Efektif11Persen() => Assert.Equal(1_100_000, TaxCalc.Ppn(10_000_000)[2].Amount, 4);
}

public class AiPlumbingTests
{
    [Fact]
    public void ParsePlan_ToleranTerhadapTeksSekitar_DanKomaBerlebih()
    {
        var raw = "<think>hmm</think>Berikut rencananya:\n```json\n{\"message\":\"ok\",\"steps\":[{\"action\":\"write_formula\",\"args\":{\"range\":\"F2:F10\",\"formula\":\"=D2*E2\",},},]}\n```";
        var plan = JsonRepair.ParsePlan(raw);
        Assert.Equal("ok", plan.Message);
        Assert.Single(plan.Steps);
        Assert.Equal("=D2*E2", plan.Steps[0].Str("formula"));
    }

    [Fact]
    public void ParsePlan_KunciBahasaIndonesia()
    {
        var plan = JsonRepair.ParsePlan("{\"pesan\":\"x\",\"langkah\":[{\"aksi\":\"insert_toc\",\"param\":{\"levels\":2}}]}");
        Assert.Equal("insert_toc", plan.Steps[0].Action);
        Assert.Equal(2, plan.Steps[0].IntOr("levels", 0));
    }

    [Fact]
    public void Validate_EnumDanArgumenWajib()
    {
        var plan = JsonRepair.ParsePlan("{\"message\":\"\",\"steps\":[{\"action\":\"FORMAT_NUMBER\",\"args\":{\"range\":\"A1\",\"format\":\"RUPIAH\"}},{\"action\":\"tidak_ada\",\"args\":{}}]}");
        var problems = ToolCatalog.Validate(OfficeApp.Excel, plan);
        Assert.Equal("format_number", plan.Steps[0].Action);
        Assert.Equal("rupiah", plan.Steps[0].Str("format"));
        Assert.False(plan.Steps[1].Enabled);
        Assert.Single(problems);
    }

    [Fact]
    public void Select_MemilihAksiRelevan()
    {
        var tools = ToolCatalog.Select(OfficeApp.Excel, "uji validitas dan reliabilitas kuesioner");
        Assert.Equal("validity_reliability", tools[0].Name);
        var w = ToolCatalog.Select(OfficeApp.Word, "buatkan daftar isi dan nomor halaman romawi");
        Assert.Contains(w, t => t.Name == "insert_toc");
        Assert.Contains(w, t => t.Name == "insert_page_numbers");
    }

    [Fact]
    public void Grammar_MemuatSemuaNamaAksi()
    {
        var g = PlanGrammar.Build(ToolCatalog.Excel);
        Assert.StartsWith("root ::=", g);
        Assert.Contains("\"\\\"write_formula\\\"\"", g);
        Assert.Contains("string ::=", g);
    }

    [Fact]
    public void NamaAksiUnik_DanDeskripsiTerisi()
    {
        foreach (var app in new[] { OfficeApp.Word, OfficeApp.Excel })
        {
            var tools = ToolCatalog.For(app);
            Assert.Equal(tools.Count, tools.Select(t => t.Name).Distinct().Count());
            Assert.All(tools, t => Assert.False(string.IsNullOrWhiteSpace(t.Description)));
        }
        Assert.True(ToolCatalog.Word.Count >= 45);
        Assert.True(ToolCatalog.Excel.Count >= 50);
    }

    [Fact]
    public void PromptPerencana_MemuatKonteksDanLarangan2016()
    {
        var ctx = new OfficeContext { App = OfficeApp.Excel, ActiveSheet = "Data", DataRegion = "A1:E10", DataRowCount = 9 };
        ctx.Headers.Add("A=Tanggal");
        var msgs = Prompts.BuildPlanMessages(OfficeApp.Excel, "total per wilayah", ctx, ToolCatalog.Select(OfficeApp.Excel, "total per wilayah"));
        Assert.Contains("XLOOKUP", msgs[0].Content);
        Assert.Contains("A=Tanggal", msgs.Last().Content);
    }
}

public class ExcelKitTests
{
    [Fact]
    public void Formula_KompatibilitasDanSintaks()
    {
        Assert.NotEmpty(FormulaTools.CompatibilityIssues("=XLOOKUP(A2,B:B,C:C)"));
        Assert.Empty(FormulaTools.CompatibilityIssues("=INDEX(C:C,MATCH(A2,B:B,0))"));
        Assert.Equal("Kurang 1 kurung tutup.", FormulaTools.SyntaxProblem("=SUM(A1:A3"));
        Assert.Equal("=SUMIFS(F:F,B:B,\"Jawa;Bali\")", FormulaTools.Normalize("SUMIFS(F:F;B:B;\"Jawa;Bali\")"));
    }

    [Fact]
    public void Formula_Penjelasan()
    {
        var text = FormulaTools.Explain("=IFERROR(VLOOKUP($A2,Data!$A$2:$D$100,3,FALSE),\"-\")");
        Assert.Contains("VLOOKUP", text);
        Assert.Contains("IFERROR", text);
        Assert.Contains("absolut", text);
    }

    [Fact]
    public void Kolom_Huruf() { Assert.Equal("AG", FormulaTools.ColumnLetter(33)); Assert.Equal(33, FormulaTools.ColumnNumber("AG")); }

    [Fact]
    public void Pembersih_Data()
    {
        var grid = new object[,] { { "  budi   santoso ", "Rp 1.500.000" }, { "12 Januari 2024", "0812-111-222" } };
        int changed = DataCleaner.CleanGrid(grid, new[] { "trim", "text_to_number", "normalize_date" });
        Assert.Equal(3, changed);
        Assert.Equal("budi santoso", grid[0, 0]);
        Assert.Equal(1500000d, grid[0, 1]);
        Assert.Equal(new DateTime(2024, 1, 12).ToOADate(), grid[1, 0]);
    }

    [Fact]
    public void Vba_DeteksiRisiko()
    {
        var risks = VbaSafety.Scan("Sub X()\nKill \"C:\\a.txt\"\nShell \"cmd\"\nEnd Sub");
        Assert.Equal(2, risks.Count);
        Assert.Equal("X", VbaSafety.FirstSubName("Sub X()\nEnd Sub"));
    }

    [Fact]
    public void Templat_Excel_PlaceholderDiganti()
    {
        Assert.Equal("=SUM(F5:F24)", ExcelTemplates.Expand("=SUM({col}{first}:{col}{last})", 25, 5, 24, "F"));
        Assert.Equal("=F25*11%", ExcelTemplates.Expand("=F{last+1}*11%", 26, 5, 24, "F"));
        Assert.NotNull(ExcelTemplates.Find("rab"));
        Assert.True(ExcelTemplates.All.Count >= 15);
    }
}

public class AcademicTests
{
    private static readonly Reference Jurnal = new()
    {
        Type = "journal", Authors = { "Suharsimi Arikunto", "Budi Santoso", "Citra Dewi" }, Year = "2020",
        Title = "Literasi digital siswa sekolah dasar", Journal = "Jurnal Pendidikan", Volume = "5", Issue = "2", Pages = "10-20", Doi = "10.1234/jp.v5i2"
    };

    [Fact]
    public void Apa7_Entri_Dan_Sitasi()
    {
        var e = CitationFormatter.Entry(Jurnal, "APA");
        Assert.Equal("Arikunto, S., Santoso, B., & Dewi, C. (2020). Literasi digital siswa sekolah dasar. *Jurnal Pendidikan, 5*(2), 10–20. https://doi.org/10.1234/jp.v5i2", e);
        Assert.Equal("(Arikunto dkk., 2020, hlm. 12)", CitationFormatter.InText(Jurnal, "APA", "12"));
    }

    [Fact]
    public void Ieee_Buku()
    {
        var book = new Reference { Type = "book", Authors = { "Sugiyono" }, Year = "2019", Title = "Metode Penelitian Pendidikan", Publisher = "Alfabeta", City = "Bandung" };
        Assert.Equal("[1] Sugiyono, *Metode Penelitian Pendidikan*. Bandung: Alfabeta, 2019.", CitationFormatter.Entry(book, "IEEE", 1));
        Assert.Equal("Sugiyono. (2019). *Metode Penelitian Pendidikan*. Alfabeta.", CitationFormatter.Entry(book, "APA7"));
    }

    [Fact]
    public void Templat_Dokumen_DiisiDanPenandaKosong()
    {
        var tpl = DocumentTemplates.Find("surat_undangan")!;
        var blocks = tpl.Render(new Dictionary<string, string> { ["acara"] = "Rapat Komite", ["kota"] = "Bandung" });
        var all = string.Join("\n", blocks.SelectMany(b => new[] { b.Text ?? "" }.Concat(b.Rows?.SelectMany(r => r) ?? Array.Empty<string>())));
        Assert.Contains("Rapat Komite", all);
        Assert.Contains("[Tempat]", all);
        Assert.True(DocumentTemplates.All.Count >= 25);
        Assert.Equal(DocumentTemplates.All.Count, DocumentTemplates.All.Select(t => t.Id).Distinct().Count());
    }

    [Fact]
    public void Blok_ParseDariJsonAi()
    {
        var blocks = DocBlock.ParseMany(Newtonsoft.Json.Linq.JArray.Parse("[{\"type\":\"h2\",\"text\":\"A\"},\"paragraf\",{\"type\":\"list\",\"items\":[\"x\",\"y\"]}]"));
        Assert.Equal("heading", blocks[0].Type);
        Assert.Equal(2, blocks[0].Level);
        Assert.Equal("paragraph", blocks[1].Type);
        Assert.Equal("bullets", blocks[2].Type);
    }
}
