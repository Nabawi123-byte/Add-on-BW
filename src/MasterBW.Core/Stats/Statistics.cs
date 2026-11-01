using System;
using System.Collections.Generic;
using System.Linq;

namespace MasterBW.Core.Stats
{
    /// <summary>Statistik deskriptif per variabel (definisi sama dengan fungsi Excel).</summary>
    public sealed class Descriptive
    {
        public int N { get; private set; }
        public double Sum { get; private set; }
        public double Mean { get; private set; }
        public double Median { get; private set; }
        public double? Mode { get; private set; }
        public double Variance { get; private set; }
        public double StdDev { get; private set; }
        public double StdError { get; private set; }
        public double Min { get; private set; }
        public double Max { get; private set; }
        public double Range => Max - Min;
        public double Q1 { get; private set; }
        public double Q3 { get; private set; }
        public double Iqr => Q3 - Q1;
        public double Skewness { get; private set; }
        public double Kurtosis { get; private set; }
        public double CoefVariation => Mean == 0 ? double.NaN : StdDev / Mean * 100;

        public static Descriptive Of(IList<double> data)
        {
            var x = data.Where(v => !double.IsNaN(v)).OrderBy(v => v).ToArray();
            var d = new Descriptive { N = x.Length };
            if (x.Length == 0) return d;
            d.Sum = x.Sum();
            d.Mean = d.Sum / x.Length;
            d.Min = x[0];
            d.Max = x[x.Length - 1];
            d.Median = Percentile(x, 0.5);
            d.Q1 = Percentile(x, 0.25);
            d.Q3 = Percentile(x, 0.75);
            var groups = x.GroupBy(v => v).Where(g => g.Count() > 1).OrderByDescending(g => g.Count()).ThenBy(g => Array.IndexOf(x, g.Key)).ToList();
            d.Mode = groups.Count > 0 ? groups[0].Key : (double?)null;
            if (x.Length > 1)
            {
                double ss = x.Sum(v => (v - d.Mean) * (v - d.Mean));
                d.Variance = ss / (x.Length - 1);
                d.StdDev = Math.Sqrt(d.Variance);
                d.StdError = d.StdDev / Math.Sqrt(x.Length);
            }
            int n = x.Length;
            if (n > 2 && d.StdDev > 0)
            {
                double s3 = x.Sum(v => Math.Pow((v - d.Mean) / d.StdDev, 3));
                d.Skewness = n / ((n - 1.0) * (n - 2.0)) * s3;
            }
            if (n > 3 && d.StdDev > 0)
            {
                double s4 = x.Sum(v => Math.Pow((v - d.Mean) / d.StdDev, 4));
                d.Kurtosis = n * (n + 1.0) / ((n - 1.0) * (n - 2.0) * (n - 3.0)) * s4 - 3.0 * (n - 1.0) * (n - 1.0) / ((n - 2.0) * (n - 3.0));
            }
            return d;
        }

        /// <summary>Persentil inklusif (sama dengan PERCENTILE.INC / QUARTILE.INC Excel). Data harus terurut.</summary>
        public static double Percentile(double[] sorted, double p)
        {
            if (sorted.Length == 0) return double.NaN;
            double h = (sorted.Length - 1) * p;
            int lo = (int)Math.Floor(h);
            int hi = Math.Min(lo + 1, sorted.Length - 1);
            return sorted[lo] + (h - lo) * (sorted[hi] - sorted[lo]);
        }
    }

    /// <summary>Uji statistik inferensial lengkap untuk kebutuhan skripsi/tesis/riset kantor.</summary>
    public static class Statistics
    {
        public const double Alpha = 0.05;

        public static string Sig(double p, double alpha = Alpha) => p < alpha ? "signifikan" : "tidak signifikan";

        // ================================================================ Deskriptif
        public static AnalysisResult DescriptiveAnalysis(IList<string> names, IList<IList<double>> columns)
        {
            var res = new AnalysisResult("Statistik Deskriptif");
            var t = new StatTable("Statistik Deskriptif", "Statistik");
            t.Headers.AddRange(names);
            var ds = columns.Select(c => Descriptive.Of(c)).ToList();
            void R(string label, Func<Descriptive, object> f) => t.Row(new object[] { label }.Concat(ds.Select(f)).ToArray());
            R("N", d => d.N);
            R("Jumlah", d => d.Sum);
            R("Rata-rata (Mean)", d => d.Mean);
            R("Galat baku rata-rata", d => d.StdError);
            R("Median", d => d.Median);
            R("Modus", d => d.Mode.HasValue ? (object)d.Mode.Value : "-");
            R("Simpangan baku (SD)", d => d.StdDev);
            R("Varians", d => d.Variance);
            R("Minimum", d => d.Min);
            R("Maksimum", d => d.Max);
            R("Rentang", d => d.Range);
            R("Kuartil 1 (Q1)", d => d.Q1);
            R("Kuartil 3 (Q3)", d => d.Q3);
            R("Koefisien variasi (%)", d => d.CoefVariation);
            R("Kemencengan (Skewness)", d => d.Skewness);
            R("Keruncingan (Kurtosis)", d => d.Kurtosis);
            res.Tables.Add(t);
            for (int i = 0; i < names.Count; i++)
            {
                var d = ds[i];
                if (d.N == 0) continue;
                var shape = Math.Abs(d.Skewness) < 0.5 ? "relatif simetris" : d.Skewness > 0 ? "menceng ke kanan (banyak nilai kecil, sedikit nilai besar)" : "menceng ke kiri";
                res.Interpretation.Add($"{names[i]}: rata-rata {d.Mean:0.##} (SD {d.StdDev:0.##}), median {d.Median:0.##}, rentang {d.Min:0.##}–{d.Max:0.##}; sebaran {shape}.");
            }
            return res;
        }

        // ================================================================ Korelasi
        public static double Pearson(IList<double> x, IList<double> y)
        {
            int n = Math.Min(x.Count, y.Count);
            double mx = x.Take(n).Average(), my = y.Take(n).Average();
            double sxy = 0, sxx = 0, syy = 0;
            for (int i = 0; i < n; i++)
            {
                double dx = x[i] - mx, dy = y[i] - my;
                sxy += dx * dy; sxx += dx * dx; syy += dy * dy;
            }
            return sxx == 0 || syy == 0 ? double.NaN : sxy / Math.Sqrt(sxx * syy);
        }

        public static double PearsonP(double r, int n)
        {
            if (n < 3 || double.IsNaN(r)) return double.NaN;
            if (Math.Abs(r) >= 1) return 0;
            double t = r * Math.Sqrt((n - 2) / (1 - r * r));
            return Distributions.TTwoTailed(t, n - 2);
        }

        /// <summary>Peringkat dengan rata-rata untuk nilai kembar (seperti RANK.AVG).</summary>
        public static double[] Ranks(IList<double> x)
        {
            var idx = Enumerable.Range(0, x.Count).OrderBy(i => x[i]).ToArray();
            var ranks = new double[x.Count];
            int k = 0;
            while (k < idx.Length)
            {
                int j = k;
                while (j + 1 < idx.Length && x[idx[j + 1]] == x[idx[k]]) j++;
                double avg = (k + j) / 2.0 + 1;
                for (int m = k; m <= j; m++) ranks[idx[m]] = avg;
                k = j + 1;
            }
            return ranks;
        }

        public static double Spearman(IList<double> x, IList<double> y) => Pearson(Ranks(x), Ranks(y));

        public static string Strength(double r)
        {
            double a = Math.Abs(r);
            if (a < 0.2) return "sangat lemah";
            if (a < 0.4) return "lemah";
            if (a < 0.6) return "sedang";
            if (a < 0.8) return "kuat";
            return "sangat kuat";
        }

        public static AnalysisResult CorrelationMatrix(IList<string> names, IList<IList<double>> cols, bool spearman = false)
        {
            var res = new AnalysisResult("Korelasi " + (spearman ? "Spearman" : "Pearson"));
            var rt = new StatTable("Koefisien korelasi (r)", "");
            var pt = new StatTable("Sig. (2-tailed)", "");
            rt.Headers.AddRange(names);
            pt.Headers.AddRange(names);
            int k = names.Count;
            for (int i = 0; i < k; i++)
            {
                var rrow = new object[k + 1];
                var prow = new object[k + 1];
                rrow[0] = prow[0] = names[i];
                for (int j = 0; j < k; j++)
                {
                    var (a, b) = Paired(cols[i], cols[j]);
                    double r = i == j ? 1 : spearman ? Spearman(a, b) : Pearson(a, b);
                    double p = i == j ? double.NaN : PearsonP(r, a.Count);
                    rrow[j + 1] = r;
                    prow[j + 1] = i == j ? (object)"-" : p;
                    if (j > i) res.Interpretation.Add($"{names[i]} – {names[j]}: r = {r:0.###} ({Strength(r)}, arah {(r >= 0 ? "positif" : "negatif")}), p = {p:0.####} → {Sig(p)}.");
                }
                rt.Rows.Add(rrow);
                pt.Rows.Add(prow);
            }
            res.Tables.Add(rt);
            res.Tables.Add(pt);
            return res;
        }

        private static (List<double>, List<double>) Paired(IList<double> a, IList<double> b)
        {
            var x = new List<double>(); var y = new List<double>();
            for (int i = 0; i < Math.Min(a.Count, b.Count); i++)
                if (!double.IsNaN(a[i]) && !double.IsNaN(b[i])) { x.Add(a[i]); y.Add(b[i]); }
            return (x, y);
        }

        // ================================================================ Regresi
        public sealed class RegressionModel
        {
            public double[] Coef;
            public double[] StdErr;
            public double[] T;
            public double[] P;
            public double[] Vif;
            public double R2, AdjR2, F, FP, See, DurbinWatson;
            public int N, K;
            public double[] Residuals;
        }

        /// <summary>Regresi linear OLS. xs[j][i] = nilai prediktor j pada observasi i.</summary>
        public static RegressionModel Ols(IList<double> y, IList<IList<double>> xs)
        {
            int k = xs.Count;
            var rows = new List<(double y, double[] x)>();
            for (int i = 0; i < y.Count; i++)
            {
                if (double.IsNaN(y[i])) continue;
                var xr = new double[k];
                bool ok = true;
                for (int j = 0; j < k; j++)
                {
                    if (i >= xs[j].Count || double.IsNaN(xs[j][i])) { ok = false; break; }
                    xr[j] = xs[j][i];
                }
                if (ok) rows.Add((y[i], xr));
            }
            int n = rows.Count, p = k + 1;
            if (n <= p) throw new ArgumentException("Jumlah data terlalu sedikit untuk regresi (n harus > jumlah variabel + 1).");
            var X = new double[n, p];
            var Y = new double[n];
            for (int i = 0; i < n; i++)
            {
                X[i, 0] = 1;
                for (int j = 0; j < k; j++) X[i, j + 1] = rows[i].x[j];
                Y[i] = rows[i].y;
            }
            var xtx = new double[p, p];
            var xty = new double[p];
            for (int a = 0; a < p; a++)
            {
                for (int b = 0; b < p; b++)
                {
                    double s = 0;
                    for (int i = 0; i < n; i++) s += X[i, a] * X[i, b];
                    xtx[a, b] = s;
                }
                double t = 0;
                for (int i = 0; i < n; i++) t += X[i, a] * Y[i];
                xty[a] = t;
            }
            var inv = Matrix.Inverse(xtx) ?? throw new ArgumentException("Matriks singular: ada variabel X yang identik/kolinear sempurna.");
            var beta = new double[p];
            for (int a = 0; a < p; a++) for (int b = 0; b < p; b++) beta[a] += inv[a, b] * xty[b];

            double ymean = Y.Average(), sse = 0, sst = 0;
            var resid = new double[n];
            for (int i = 0; i < n; i++)
            {
                double fit = 0;
                for (int a = 0; a < p; a++) fit += X[i, a] * beta[a];
                resid[i] = Y[i] - fit;
                sse += resid[i] * resid[i];
                sst += (Y[i] - ymean) * (Y[i] - ymean);
            }
            int dfe = n - p;
            double mse = sse / dfe;
            var m = new RegressionModel
            {
                N = n, K = k, Coef = beta, Residuals = resid,
                StdErr = new double[p], T = new double[p], P = new double[p], Vif = new double[k],
                R2 = sst == 0 ? 0 : 1 - sse / sst,
                See = Math.Sqrt(mse)
            };
            m.AdjR2 = 1 - (1 - m.R2) * (n - 1) / dfe;
            m.F = k == 0 || m.R2 >= 1 ? double.PositiveInfinity : (m.R2 / k) / ((1 - m.R2) / dfe);
            m.FP = double.IsInfinity(m.F) ? 0 : Distributions.FSurvival(m.F, k, dfe);
            for (int a = 0; a < p; a++)
            {
                m.StdErr[a] = Math.Sqrt(mse * inv[a, a]);
                m.T[a] = m.StdErr[a] == 0 ? double.NaN : beta[a] / m.StdErr[a];
                m.P[a] = Distributions.TTwoTailed(m.T[a], dfe);
            }
            double dwNum = 0;
            for (int i = 1; i < n; i++) dwNum += (resid[i] - resid[i - 1]) * (resid[i] - resid[i - 1]);
            m.DurbinWatson = sse == 0 ? double.NaN : dwNum / sse;
            for (int j = 0; j < k; j++)
            {
                if (k == 1) { m.Vif[j] = 1; continue; }
                var others = Enumerable.Range(0, k).Where(o => o != j).Select(o => (IList<double>)rows.Select(r => r.x[o]).ToList()).ToList();
                try
                {
                    var aux = Ols(rows.Select(r => r.x[j]).ToList(), others);
                    m.Vif[j] = aux.R2 >= 1 ? double.PositiveInfinity : 1 / (1 - aux.R2);
                }
                catch (ArgumentException) { m.Vif[j] = double.PositiveInfinity; }
            }
            return m;
        }

        public static AnalysisResult RegressionAnalysis(string yName, IList<double> y, IList<string> xNames, IList<IList<double>> xs)
        {
            var m = Ols(y, xs);
            var res = new AnalysisResult("Regresi Linear " + (xs.Count > 1 ? "Berganda" : "Sederhana"));
            res.Tables.Add(new StatTable("Ringkasan Model", "R", "R Square", "Adjusted R Square", "Std. Error of Estimate", "Durbin-Watson", "N")
                .Row(Math.Sqrt(Math.Max(0, m.R2)), m.R2, m.AdjR2, m.See, m.DurbinWatson, m.N));
            res.Tables.Add(new StatTable("ANOVA (Uji F)", "F hitung", "df1", "df2", "Sig.")
                .Row(m.F, m.K, m.N - m.K - 1, m.FP));
            var ct = new StatTable("Koefisien (Uji t)", "Variabel", "B", "Std. Error", "t hitung", "Sig.", "VIF");
            ct.Row("(Konstanta)", m.Coef[0], m.StdErr[0], m.T[0], m.P[0], "-");
            for (int j = 0; j < xs.Count; j++) ct.Row(xNames[j], m.Coef[j + 1], m.StdErr[j + 1], m.T[j + 1], m.P[j + 1], m.Vif[j]);
            res.Tables.Add(ct);

            var eq = yName + " = " + m.Coef[0].ToString("0.####") + string.Concat(xNames.Select((n, j) => (m.Coef[j + 1] >= 0 ? " + " : " − ") + Math.Abs(m.Coef[j + 1]).ToString("0.####") + "·" + n));
            res.Interpretation.Add("Persamaan regresi: " + eq);
            res.Interpretation.Add($"R² = {m.R2:0.####}: {m.R2 * 100:0.##}% variasi {yName} dijelaskan oleh model; sisanya {(1 - m.R2) * 100:0.##}% oleh faktor lain.");
            res.Interpretation.Add($"Uji F: F = {m.F:0.###}, Sig. = {m.FP:0.####} → secara simultan {Sig(m.FP)}.");
            for (int j = 0; j < xs.Count; j++)
                res.Interpretation.Add($"{xNames[j]}: B = {m.Coef[j + 1]:0.####}, t = {m.T[j + 1]:0.###}, Sig. = {m.P[j + 1]:0.####} → pengaruh parsial {Sig(m.P[j + 1])} ({(m.Coef[j + 1] >= 0 ? "positif" : "negatif")}).");
            if (xs.Count > 1)
            {
                var multi = xNames.Where((n, j) => m.Vif[j] >= 10).ToList();
                res.Interpretation.Add(multi.Count == 0 ? "Multikolinearitas: semua VIF < 10 (tidak terjadi multikolinearitas)." : "Multikolinearitas terindikasi pada: " + string.Join(", ", multi) + " (VIF ≥ 10).");
            }
            res.Interpretation.Add($"Durbin-Watson = {m.DurbinWatson:0.###} (nilai mendekati 2 menunjukkan tidak ada autokorelasi).");
            var norm = Normality(m.Residuals);
            res.Interpretation.Add($"Normalitas residual (Shapiro-Wilk): W = {norm.W:0.####}, Sig. = {norm.Wp:0.####} → residual {(norm.Wp >= Alpha ? "berdistribusi normal" : "tidak normal")}.");
            return res;
        }

        // ================================================================ Uji t
        public static AnalysisResult TTest(string type, IList<double> a, IList<double> b = null, double mu = 0, string nameA = "Kelompok 1", string nameB = "Kelompok 2")
        {
            var res = new AnalysisResult("Uji t");
            var da = Descriptive.Of(a);
            switch ((type ?? "").ToLowerInvariant())
            {
                case "one_sample":
                {
                    double t = (da.Mean - mu) / da.StdError;
                    double df = da.N - 1, p = Distributions.TTwoTailed(t, df);
                    res.Title = "Uji t Satu Sampel";
                    res.Tables.Add(new StatTable("Uji t Satu Sampel (nilai uji = " + mu + ")", "Variabel", "N", "Mean", "SD", "t hitung", "df", "Sig. (2-tailed)", "Selisih rata-rata")
                        .Row(nameA, da.N, da.Mean, da.StdDev, t, df, p, da.Mean - mu));
                    res.Interpretation.Add($"Rata-rata {nameA} ({da.Mean:0.##}) {(p < Alpha ? "berbeda signifikan" : "tidak berbeda signifikan")} dari {mu} (t = {t:0.###}, Sig. = {p:0.####}).");
                    break;
                }
                case "paired":
                {
                    var (x, y) = Paired(a, b);
                    var diff = x.Zip(y, (p1, p2) => p1 - p2).ToList();
                    var dd = Descriptive.Of(diff);
                    double t = dd.Mean / dd.StdError, df = dd.N - 1, p = Distributions.TTwoTailed(t, df);
                    double tcrit = Distributions.TInvTwoTailed(Alpha, df);
                    var r = Pearson(x, y);
                    res.Title = "Uji t Sampel Berpasangan";
                    res.Tables.Add(new StatTable("Statistik Sampel Berpasangan", "Variabel", "N", "Mean", "SD")
                        .Row(nameA, x.Count, x.Average(), Descriptive.Of(x).StdDev).Row(nameB, y.Count, y.Average(), Descriptive.Of(y).StdDev));
                    res.Tables.Add(new StatTable("Uji t Berpasangan", "Pasangan", "Mean selisih", "SD selisih", "t hitung", "t tabel", "df", "Sig. (2-tailed)", "Korelasi")
                        .Row(nameA + " − " + nameB, dd.Mean, dd.StdDev, t, tcrit, df, p, r));
                    res.Interpretation.Add($"Terdapat perbedaan {(p < Alpha ? "yang signifikan" : "yang tidak signifikan")} antara {nameA} dan {nameB} (t = {t:0.###}, Sig. = {p:0.####}). Rata-rata selisih = {dd.Mean:0.###}.");
                    break;
                }
                default:
                {
                    var db = Descriptive.Of(b);
                    double n1 = da.N, n2 = db.N, v1 = da.Variance, v2 = db.Variance;
                    double sp = Math.Sqrt(((n1 - 1) * v1 + (n2 - 1) * v2) / (n1 + n2 - 2));
                    double tPooled = (da.Mean - db.Mean) / (sp * Math.Sqrt(1 / n1 + 1 / n2));
                    double dfPooled = n1 + n2 - 2, pPooled = Distributions.TTwoTailed(tPooled, dfPooled);
                    double se = Math.Sqrt(v1 / n1 + v2 / n2);
                    double tWelch = (da.Mean - db.Mean) / se;
                    double dfWelch = Math.Pow(v1 / n1 + v2 / n2, 2) / (Math.Pow(v1 / n1, 2) / (n1 - 1) + Math.Pow(v2 / n2, 2) / (n2 - 1));
                    double pWelch = Distributions.TTwoTailed(tWelch, dfWelch);
                    var lev = Levene(new[] { a, b });
                    bool equalVar = lev.p >= Alpha;
                    double cohenD = (da.Mean - db.Mean) / sp;
                    res.Title = "Uji t Sampel Independen";
                    res.Tables.Add(new StatTable("Statistik Kelompok", "Kelompok", "N", "Mean", "SD", "SE Mean")
                        .Row(nameA, da.N, da.Mean, da.StdDev, da.StdError).Row(nameB, db.N, db.Mean, db.StdDev, db.StdError));
                    res.Tables.Add(new StatTable("Uji Independent Samples t", "Asumsi", "Levene F", "Sig. Levene", "t hitung", "df", "Sig. (2-tailed)", "Selisih rata-rata")
                        .Row("Varians sama", lev.f, lev.p, tPooled, dfPooled, pPooled, da.Mean - db.Mean)
                        .Row("Varians tidak sama (Welch)", "", "", tWelch, dfWelch, pWelch, da.Mean - db.Mean));
                    double pUsed = equalVar ? pPooled : pWelch;
                    res.Interpretation.Add($"Uji Levene Sig. = {lev.p:0.####} → varians {(equalVar ? "homogen, baca baris 'Varians sama'" : "tidak homogen, baca baris Welch")}.");
                    res.Interpretation.Add($"Rata-rata {nameA} ({da.Mean:0.##}) dan {nameB} ({db.Mean:0.##}) {(pUsed < Alpha ? "berbeda signifikan" : "tidak berbeda signifikan")} (Sig. = {pUsed:0.####}). Effect size Cohen's d = {cohenD:0.###}.");
                    break;
                }
            }
            return res;
        }

        /// <summary>Uji Levene (berbasis rata-rata) untuk homogenitas varians.</summary>
        public static (double f, double p) Levene(IList<IList<double>> groups)
        {
            var dev = groups.Select(g => { var clean = g.Where(v => !double.IsNaN(v)).ToList(); var m = clean.Average(); return (IList<double>)clean.Select(v => Math.Abs(v - m)).ToList(); }).ToList();
            var a = OneWay(dev);
            return (a.F, a.P);
        }

        // ================================================================ ANOVA
        public sealed class AnovaResult { public double Ssb, Ssw, F, P; public int Dfb, Dfw; }

        public static AnovaResult OneWay(IList<IList<double>> groups)
        {
            var g = groups.Select(x => x.Where(v => !double.IsNaN(v)).ToList()).Where(x => x.Count > 0).ToList();
            double grand = g.SelectMany(x => x).Average();
            double ssb = g.Sum(x => x.Count * Math.Pow(x.Average() - grand, 2));
            double ssw = g.Sum(x => { var m = x.Average(); return x.Sum(v => (v - m) * (v - m)); });
            int dfb = g.Count - 1, dfw = g.Sum(x => x.Count) - g.Count;
            double f = dfb == 0 || dfw == 0 || ssw == 0 ? double.NaN : (ssb / dfb) / (ssw / dfw);
            return new AnovaResult { Ssb = ssb, Ssw = ssw, Dfb = dfb, Dfw = dfw, F = f, P = double.IsNaN(f) ? double.NaN : Distributions.FSurvival(f, dfb, dfw) };
        }

        public static AnalysisResult AnovaAnalysis(IList<string> names, IList<IList<double>> groups)
        {
            var a = OneWay(groups);
            var res = new AnalysisResult("ANOVA Satu Jalur");
            var desc = new StatTable("Deskripsi Kelompok", "Kelompok", "N", "Mean", "SD");
            for (int i = 0; i < names.Count; i++) { var d = Descriptive.Of(groups[i]); desc.Row(names[i], d.N, d.Mean, d.StdDev); }
            res.Tables.Add(desc);
            var lev = Levene(groups);
            res.Tables.Add(new StatTable("Uji Homogenitas (Levene)", "Levene F", "Sig.").Row(lev.f, lev.p));
            res.Tables.Add(new StatTable("ANOVA", "Sumber", "Jumlah Kuadrat", "df", "Rata-rata Kuadrat", "F hitung", "Sig.")
                .Row("Antar kelompok", a.Ssb, a.Dfb, a.Ssb / a.Dfb, a.F, a.P)
                .Row("Dalam kelompok", a.Ssw, a.Dfw, a.Ssw / a.Dfw, "", "")
                .Row("Total", a.Ssb + a.Ssw, a.Dfb + a.Dfw, "", "", ""));
            double eta = a.Ssb / (a.Ssb + a.Ssw);
            res.Interpretation.Add($"F = {a.F:0.###}, Sig. = {a.P:0.####} → terdapat perbedaan rata-rata yang {(a.P < Alpha ? "signifikan" : "tidak signifikan")} antar kelompok. Eta² = {eta:0.###}.");
            res.Interpretation.Add($"Homogenitas varians (Levene) Sig. = {lev.p:0.####} → {(lev.p >= Alpha ? "homogen" : "tidak homogen; pertimbangkan uji Welch/Kruskal-Wallis")}.");
            return res;
        }

        // ================================================================ Chi-square
        public static AnalysisResult ChiSquare(IList<string> rowNames, IList<string> colNames, double[,] observed)
        {
            int r = observed.GetLength(0), c = observed.GetLength(1);
            double total = 0;
            var rowSum = new double[r]; var colSum = new double[c];
            for (int i = 0; i < r; i++) for (int j = 0; j < c; j++) { rowSum[i] += observed[i, j]; colSum[j] += observed[i, j]; total += observed[i, j]; }
            double chi = 0; int lowExpected = 0;
            var exp = new StatTable("Frekuensi Harapan", "");
            exp.Headers.AddRange(colNames);
            for (int i = 0; i < r; i++)
            {
                var row = new object[c + 1]; row[0] = rowNames[i];
                for (int j = 0; j < c; j++)
                {
                    double e = rowSum[i] * colSum[j] / total;
                    if (e < 5) lowExpected++;
                    chi += e == 0 ? 0 : Math.Pow(observed[i, j] - e, 2) / e;
                    row[j + 1] = e;
                }
                exp.Rows.Add(row);
            }
            int df = (r - 1) * (c - 1);
            double p = Distributions.ChiSquareSurvival(chi, df);
            double v = Math.Sqrt(chi / (total * Math.Max(1, Math.Min(r, c) - 1)));
            var res = new AnalysisResult("Uji Chi-Square Independensi");
            res.Tables.Add(exp);
            res.Tables.Add(new StatTable("Uji Chi-Square", "Chi-square hitung", "df", "Sig.", "Cramer's V", "N").Row(chi, df, p, v, total));
            res.Interpretation.Add($"χ² = {chi:0.###}, df = {df}, Sig. = {p:0.####} → hubungan kedua variabel {Sig(p)}. Kekuatan hubungan (Cramer's V) = {v:0.###}.");
            if (lowExpected > 0) res.Interpretation.Add($"Perhatian: {lowExpected} sel memiliki frekuensi harapan < 5; pertimbangkan menggabungkan kategori atau uji Fisher.");
            return res;
        }

        // ================================================================ Normalitas
        public sealed class NormalityResult { public int N; public double KsD, KsP, W, Wp, Jb, JbP; }

        public static NormalityResult Normality(IList<double> data)
        {
            var x = data.Where(v => !double.IsNaN(v)).OrderBy(v => v).ToArray();
            var r = new NormalityResult { N = x.Length };
            if (x.Length < 3) { r.KsP = r.Wp = r.JbP = double.NaN; return r; }
            var d = Descriptive.Of(x);
            int n = x.Length;
            // Kolmogorov-Smirnov dengan koreksi Lilliefors
            double dmax = 0;
            for (int i = 0; i < n; i++)
            {
                double f = Distributions.NormalCdf((x[i] - d.Mean) / d.StdDev);
                dmax = Math.Max(dmax, Math.Max((i + 1.0) / n - f, f - (double)i / n));
            }
            r.KsD = dmax;
            r.KsP = LillieforsP(dmax, n);
            (r.W, r.Wp) = ShapiroWilk(x);
            double jb = n / 6.0 * (d.Skewness * d.Skewness + d.Kurtosis * d.Kurtosis / 4);
            r.Jb = jb;
            r.JbP = Distributions.ChiSquareSurvival(jb, 2);
            return r;
        }

        /// <summary>Nilai p Lilliefors (aproksimasi Dallal-Wilkinson 1986 + Stephens), seperti paket nortest R.</summary>
        public static double LillieforsP(double D, int n)
        {
            double Kd = D, nd = n;
            if (n > 100) { Kd = D * Math.Pow(n / 100.0, 0.49); nd = 100; }
            double p = Math.Exp(-7.01256 * Kd * Kd * (nd + 2.78019) + 2.99587 * Kd * Math.Sqrt(nd + 2.78019) - 0.122119 + 0.974598 / Math.Sqrt(nd) + 1.67997 / nd);
            if (p > 0.1)
            {
                double KK = (Math.Sqrt(n) - 0.01 + 0.85 / Math.Sqrt(n)) * D;
                if (KK <= 0.302) p = 1;
                else if (KK <= 0.5) p = 2.76773 - 19.828315 * KK + 80.709644 * KK * KK - 138.55152 * Math.Pow(KK, 3) + 81.218052 * Math.Pow(KK, 4);
                else if (KK <= 0.9) p = -4.901232 + 40.662806 * KK - 97.490286 * KK * KK + 94.029866 * Math.Pow(KK, 3) - 32.355711 * Math.Pow(KK, 4);
                else if (KK <= 1.31) p = 6.198765 - 19.558097 * KK + 23.186922 * KK * KK - 12.234627 * Math.Pow(KK, 3) + 2.423045 * Math.Pow(KK, 4);
                else p = 0;
            }
            return Math.Max(0, Math.Min(1, p));
        }

        /// <summary>Uji Shapiro-Wilk (algoritma Royston 1992/1995), n = 3..5000. Data harus terurut.</summary>
        public static (double W, double p) ShapiroWilk(double[] x)
        {
            int n = x.Length;
            if (n < 3) return (double.NaN, double.NaN);
            double mean = x.Average();
            double ssq = x.Sum(v => (v - mean) * (v - mean));
            if (ssq == 0) return (1, 1);
            var a = new double[n];
            if (n == 3)
            {
                a[0] = -Math.Sqrt(0.5); a[1] = 0; a[2] = Math.Sqrt(0.5);
            }
            else
            {
                var m = new double[n];
                for (int i = 0; i < n; i++) m[i] = Distributions.NormalInv((i + 1 - 0.375) / (n + 0.25));
                double mm = m.Sum(v => v * v);
                double u = 1 / Math.Sqrt(n);
                double an = m[n - 1] / Math.Sqrt(mm) + 0.221157 * u - 0.147981 * u * u - 2.071190 * Math.Pow(u, 3) + 4.434685 * Math.Pow(u, 4) - 2.706056 * Math.Pow(u, 5);
                if (n > 5)
                {
                    double an1 = m[n - 2] / Math.Sqrt(mm) + 0.042981 * u - 0.293762 * u * u - 1.752461 * Math.Pow(u, 3) + 5.682633 * Math.Pow(u, 4) - 3.582633 * Math.Pow(u, 5);
                    double phi = (mm - 2 * m[n - 1] * m[n - 1] - 2 * m[n - 2] * m[n - 2]) / (1 - 2 * an * an - 2 * an1 * an1);
                    for (int i = 2; i < n - 2; i++) a[i] = m[i] / Math.Sqrt(phi);
                    a[n - 1] = an; a[n - 2] = an1; a[0] = -an; a[1] = -an1;
                }
                else
                {
                    double phi = (mm - 2 * m[n - 1] * m[n - 1]) / (1 - 2 * an * an);
                    for (int i = 1; i < n - 1; i++) a[i] = m[i] / Math.Sqrt(phi);
                    a[n - 1] = an; a[0] = -an;
                }
            }
            double b = 0;
            for (int i = 0; i < n; i++) b += a[i] * x[i];
            double W = Math.Min(1, b * b / ssq);
            double p;
            if (n == 3)
            {
                p = 6 / Math.PI * (Math.Asin(Math.Sqrt(W)) - Math.Asin(Math.Sqrt(0.75)));
            }
            else if (n <= 11)
            {
                double gamma = 0.459 * n - 2.273;
                double mu = 0.5440 - 0.39978 * n + 0.025054 * n * n - 0.0006714 * Math.Pow(n, 3);
                double sigma = Math.Exp(1.3822 - 0.77857 * n + 0.062767 * n * n - 0.0020322 * Math.Pow(n, 3));
                double w1 = gamma - Math.Log(1 - W);
                double z = w1 <= 0 ? 10 : (-Math.Log(w1) - mu) / sigma;
                p = 1 - Distributions.NormalCdf(z);
            }
            else
            {
                double ln = Math.Log(n);
                double mu = -1.5861 - 0.31082 * ln - 0.083751 * ln * ln + 0.0038915 * Math.Pow(ln, 3);
                double sigma = Math.Exp(-0.4803 - 0.082676 * ln + 0.0030302 * ln * ln);
                double z = (Math.Log(1 - W) - mu) / sigma;
                p = 1 - Distributions.NormalCdf(z);
            }
            return (W, Math.Max(0, Math.Min(1, p)));
        }

        public static AnalysisResult NormalityAnalysis(IList<string> names, IList<IList<double>> cols)
        {
            var res = new AnalysisResult("Uji Normalitas");
            var t = new StatTable("Uji Normalitas", "Variabel", "N", "Kolmogorov-Smirnov D", "Sig. KS (Lilliefors)", "Shapiro-Wilk W", "Sig. SW", "Jarque-Bera", "Sig. JB", "Kesimpulan");
            for (int i = 0; i < names.Count; i++)
            {
                var r = Normality(cols[i]);
                double pMain = r.N <= 50 ? r.Wp : r.KsP;
                var concl = pMain >= Alpha ? "Normal" : "Tidak normal";
                t.Row(names[i], r.N, r.KsD, r.KsP, r.W, r.Wp, r.Jb, r.JbP, concl);
                res.Interpretation.Add($"{names[i]}: Sig. {(r.N <= 50 ? "Shapiro-Wilk" : "Kolmogorov-Smirnov")} = {pMain:0.####} {(pMain >= Alpha ? "≥" : "<")} 0,05 → data {concl.ToLowerInvariant()}.");
            }
            t.Note("Untuk n ≤ 50 acuan utama Shapiro-Wilk; n > 50 acuan Kolmogorov-Smirnov (Lilliefors).");
            res.Tables.Add(t);
            return res;
        }

        // ================================================================ Validitas & reliabilitas
        /// <summary>items[j][i] = skor item j responden i.</summary>
        public static AnalysisResult ValidityReliability(IList<string> itemNames, IList<IList<double>> items, double alpha = 0.05)
        {
            int k = items.Count;
            int n = items.Min(c => c.Count);
            var total = new double[n];
            for (int i = 0; i < n; i++) for (int j = 0; j < k; j++) total[i] += items[j][i];
            double rTable = Distributions.RTable(n, alpha);
            var res = new AnalysisResult("Uji Validitas dan Reliabilitas");
            var vt = new StatTable("Uji Validitas Item (Korelasi Item-Total)", "Item", "r hitung", "r tabel", "Sig.", "r item-total terkoreksi", "Alpha jika item dihapus", "Keterangan");
            double sumItemVar = items.Sum(c => Descriptive.Of(c.Take(n).ToList()).Variance);
            double totalVar = Descriptive.Of(total).Variance;
            double cronbach = k < 2 || totalVar == 0 ? double.NaN : k / (k - 1.0) * (1 - sumItemVar / totalVar);
            int valid = 0;
            for (int j = 0; j < k; j++)
            {
                var item = items[j].Take(n).ToList();
                double r = Pearson(item, total);
                double p = PearsonP(r, n);
                var rest = total.Select((tv, i) => tv - item[i]).ToList();
                double rc = Pearson(item, rest);
                double varRest = Descriptive.Of(rest).Variance;
                double alphaDel = k < 3 || varRest == 0 ? double.NaN : (k - 1) / (k - 2.0) * (1 - (sumItemVar - Descriptive.Of(item).Variance) / varRest);
                bool ok = r > rTable;
                if (ok) valid++;
                vt.Row(itemNames[j], r, rTable, p, rc, alphaDel, ok ? "Valid" : "Tidak valid");
            }
            vt.Note($"r tabel untuk n = {n}, df = {n - 2}, α = {alpha} (dua sisi) = {rTable:0.####}. Item valid jika r hitung > r tabel.");
            res.Tables.Add(vt);
            res.Tables.Add(new StatTable("Uji Reliabilitas", "Cronbach's Alpha", "Jumlah item", "N responden", "Kategori")
                .Row(cronbach, k, n, AlphaCategory(cronbach)));
            res.Interpretation.Add($"{valid} dari {k} item valid (r hitung > r tabel {rTable:0.###}).");
            res.Interpretation.Add($"Cronbach's Alpha = {cronbach:0.###} ({AlphaCategory(cronbach)}); instrumen {(cronbach >= 0.6 ? "reliabel" : "belum reliabel")} (batas umum ≥ 0,60; ketat ≥ 0,70).");
            return res;
        }

        public static string AlphaCategory(double a)
        {
            if (double.IsNaN(a)) return "-";
            if (a >= 0.9) return "Sangat tinggi";
            if (a >= 0.7) return "Tinggi";
            if (a >= 0.6) return "Cukup";
            if (a >= 0.5) return "Rendah";
            return "Sangat rendah";
        }

        // ================================================================ Outlier & peramalan
        public static List<int> Outliers(IList<double> x, string method, out double lower, out double upper)
        {
            var d = Descriptive.Of(x);
            if ((method ?? "iqr").ToLowerInvariant() == "zscore")
            {
                lower = d.Mean - 3 * d.StdDev; upper = d.Mean + 3 * d.StdDev;
            }
            else
            {
                lower = d.Q1 - 1.5 * d.Iqr; upper = d.Q3 + 1.5 * d.Iqr;
            }
            double lo = lower, hi = upper;
            return Enumerable.Range(0, x.Count).Where(i => !double.IsNaN(x[i]) && (x[i] < lo || x[i] > hi)).ToList();
        }

        public sealed class ForecastResult
        {
            public List<double> Fitted = new List<double>();
            public List<double> Forecast = new List<double>();
            public string Method;
            public string Equation;
            public double Mape;
        }

        public static ForecastResult ForecastSeries(IList<double> y, int periods, string method = "linear", int window = 3)
        {
            var data = y.Where(v => !double.IsNaN(v)).ToList();
            int n = data.Count;
            if (n < 2) throw new ArgumentException("Butuh minimal 2 data historis.");
            var r = new ForecastResult { Method = method };
            switch ((method ?? "linear").ToLowerInvariant())
            {
                case "moving_average":
                {
                    window = Math.Max(1, Math.Min(window, n));
                    for (int i = 0; i < n; i++) r.Fitted.Add(i < window ? double.NaN : data.Skip(i - window).Take(window).Average());
                    var ext = new List<double>(data);
                    for (int h = 0; h < periods; h++) { var f = ext.Skip(ext.Count - window).Average(); r.Forecast.Add(f); ext.Add(f); }
                    r.Equation = $"Rata-rata {window} periode terakhir";
                    break;
                }
                case "exponential":
                {
                    double bestA = 0.3, bestSse = double.MaxValue;
                    for (double a = 0.05; a <= 0.951; a += 0.05)
                    {
                        double s = data[0], sse = 0;
                        for (int i = 1; i < n; i++) { sse += Math.Pow(data[i] - s, 2); s = a * data[i] + (1 - a) * s; }
                        if (sse < bestSse) { bestSse = sse; bestA = a; }
                    }
                    double level = data[0];
                    r.Fitted.Add(double.NaN);
                    for (int i = 1; i < n; i++) { r.Fitted.Add(level); level = bestA * data[i] + (1 - bestA) * level; }
                    for (int h = 0; h < periods; h++) r.Forecast.Add(level);
                    r.Equation = $"Pemulusan eksponensial, α = {bestA:0.00}";
                    break;
                }
                default:
                {
                    var t = Enumerable.Range(1, n).Select(v => (double)v).ToList();
                    double mt = t.Average(), my = data.Average();
                    double b = t.Zip(data, (ti, yi) => (ti - mt) * (yi - my)).Sum() / t.Sum(ti => (ti - mt) * (ti - mt));
                    double a = my - b * mt;
                    foreach (var ti in t) r.Fitted.Add(a + b * ti);
                    for (int h = 1; h <= periods; h++) r.Forecast.Add(a + b * (n + h));
                    r.Equation = $"Y = {a:0.###} + {b:0.###}·t";
                    break;
                }
            }
            var errs = data.Select((v, i) => (v, f: r.Fitted[i])).Where(z => !double.IsNaN(z.f) && z.v != 0).Select(z => Math.Abs((z.v - z.f) / z.v)).ToList();
            r.Mape = errs.Count == 0 ? double.NaN : errs.Average() * 100;
            return r;
        }
    }

    internal static class Matrix
    {
        /// <summary>Invers matriks dengan eliminasi Gauss-Jordan berpivot. null jika singular.</summary>
        public static double[,] Inverse(double[,] m)
        {
            int n = m.GetLength(0);
            var a = new double[n, 2 * n];
            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < n; j++) a[i, j] = m[i, j];
                a[i, n + i] = 1;
            }
            for (int col = 0; col < n; col++)
            {
                int pivot = col;
                for (int r = col + 1; r < n; r++) if (Math.Abs(a[r, col]) > Math.Abs(a[pivot, col])) pivot = r;
                if (Math.Abs(a[pivot, col]) < 1e-12) return null;
                if (pivot != col) for (int j = 0; j < 2 * n; j++) { var tmp = a[col, j]; a[col, j] = a[pivot, j]; a[pivot, j] = tmp; }
                double div = a[col, col];
                for (int j = 0; j < 2 * n; j++) a[col, j] /= div;
                for (int r = 0; r < n; r++)
                {
                    if (r == col) continue;
                    double f = a[r, col];
                    if (f == 0) continue;
                    for (int j = 0; j < 2 * n; j++) a[r, j] -= f * a[col, j];
                }
            }
            var inv = new double[n, n];
            for (int i = 0; i < n; i++) for (int j = 0; j < n; j++) inv[i, j] = a[i, n + j];
            return inv;
        }
    }
}
