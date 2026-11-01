using System;

namespace MasterBW.Core.Stats
{
    /// <summary>Fungsi khusus &amp; distribusi probabilitas (tanpa pustaka luar) untuk nilai p uji statistik.</summary>
    public static class Distributions
    {
        private static readonly double[] Lanczos =
        {
            676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.61502916214059,
            12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7
        };

        public static double LogGamma(double x)
        {
            if (x < 0.5) return Math.Log(Math.PI / Math.Abs(Math.Sin(Math.PI * x))) - LogGamma(1 - x);
            x -= 1;
            double a = 0.99999999999980993;
            double t = x + 7.5;
            for (int i = 0; i < 8; i++) a += Lanczos[i] / (x + i + 1);
            return 0.5 * Math.Log(2 * Math.PI) + (x + 0.5) * Math.Log(t) - t + Math.Log(a);
        }

        /// <summary>Fungsi beta tak lengkap teregularisasi I_x(a,b).</summary>
        public static double IncompleteBeta(double x, double a, double b)
        {
            if (x <= 0) return 0;
            if (x >= 1) return 1;
            double bt = Math.Exp(LogGamma(a + b) - LogGamma(a) - LogGamma(b) + a * Math.Log(x) + b * Math.Log(1 - x));
            if (x < (a + 1) / (a + b + 2)) return bt * BetaCf(x, a, b) / a;
            return 1 - bt * BetaCf(1 - x, b, a) / b;
        }

        private static double BetaCf(double x, double a, double b)
        {
            const int maxIter = 500;
            const double eps = 1e-15, fpmin = 1e-300;
            double qab = a + b, qap = a + 1, qam = a - 1;
            double c = 1, d = 1 - qab * x / qap;
            if (Math.Abs(d) < fpmin) d = fpmin;
            d = 1 / d;
            double h = d;
            for (int m = 1; m <= maxIter; m++)
            {
                int m2 = 2 * m;
                double aa = m * (b - m) * x / ((qam + m2) * (a + m2));
                d = 1 + aa * d; if (Math.Abs(d) < fpmin) d = fpmin;
                c = 1 + aa / c; if (Math.Abs(c) < fpmin) c = fpmin;
                d = 1 / d; h *= d * c;
                aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2));
                d = 1 + aa * d; if (Math.Abs(d) < fpmin) d = fpmin;
                c = 1 + aa / c; if (Math.Abs(c) < fpmin) c = fpmin;
                d = 1 / d;
                double del = d * c;
                h *= del;
                if (Math.Abs(del - 1) < eps) break;
            }
            return h;
        }

        /// <summary>Fungsi gamma tak lengkap bawah teregularisasi P(a,x).</summary>
        public static double GammaP(double a, double x)
        {
            if (x <= 0) return 0;
            if (x < a + 1)
            {
                double sum = 1 / a, del = sum, ap = a;
                for (int n = 0; n < 1000; n++)
                {
                    ap += 1; del *= x / ap; sum += del;
                    if (Math.Abs(del) < Math.Abs(sum) * 1e-15) break;
                }
                return sum * Math.Exp(-x + a * Math.Log(x) - LogGamma(a));
            }
            return 1 - GammaQcf(a, x);
        }

        private static double GammaQcf(double a, double x)
        {
            const double fpmin = 1e-300;
            double b = x + 1 - a, c = 1 / fpmin, d = 1 / b, h = d;
            for (int i = 1; i < 1000; i++)
            {
                double an = -i * (i - a);
                b += 2;
                d = an * d + b; if (Math.Abs(d) < fpmin) d = fpmin;
                c = b + an / c; if (Math.Abs(c) < fpmin) c = fpmin;
                d = 1 / d;
                double del = d * c;
                h *= del;
                if (Math.Abs(del - 1) < 1e-15) break;
            }
            return Math.Exp(-x + a * Math.Log(x) - LogGamma(a)) * h;
        }

        public static double Erf(double x)
        {
            // dari GammaP: erf(x) = P(0.5, x^2)
            double v = GammaP(0.5, x * x);
            return x >= 0 ? v : -v;
        }

        public static double NormalCdf(double z) => 0.5 * (1 + Erf(z / Math.Sqrt(2)));

        /// <summary>Invers CDF normal baku (algoritma Acklam + satu langkah Newton).</summary>
        public static double NormalInv(double p)
        {
            if (p <= 0) return double.NegativeInfinity;
            if (p >= 1) return double.PositiveInfinity;
            double[] a = { -3.969683028665376e+01, 2.209460984245205e+02, -2.759285104469687e+02, 1.383577518672690e+02, -3.066479806614716e+01, 2.506628277459239e+00 };
            double[] b = { -5.447609879822406e+01, 1.615858368580409e+02, -1.556989798598866e+02, 6.680131188771972e+01, -1.328068155288572e+01 };
            double[] c = { -7.784894002430293e-03, -3.223964580411365e-01, -2.400758277161838e+00, -2.549732539343734e+00, 4.374664141464968e+00, 2.938163982698783e+00 };
            double[] d = { 7.784695709041462e-03, 3.224671290700398e-01, 2.445134137142996e+00, 3.754408661907416e+00 };
            double q, r, x;
            if (p < 0.02425)
            {
                q = Math.Sqrt(-2 * Math.Log(p));
                x = (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
            }
            else if (p > 1 - 0.02425)
            {
                q = Math.Sqrt(-2 * Math.Log(1 - p));
                x = -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
            }
            else
            {
                q = p - 0.5; r = q * q;
                x = (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1);
            }
            double e = NormalCdf(x) - p;
            double u = e * Math.Sqrt(2 * Math.PI) * Math.Exp(x * x / 2);
            return x - u / (1 + x * u / 2);
        }

        /// <summary>CDF distribusi t Student.</summary>
        public static double TCdf(double t, double df)
        {
            double x = df / (df + t * t);
            double tail = 0.5 * IncompleteBeta(x, df / 2, 0.5);
            return t >= 0 ? 1 - tail : tail;
        }

        /// <summary>Nilai p dua sisi untuk statistik t.</summary>
        public static double TTwoTailed(double t, double df) => IncompleteBeta(df / (df + t * t), df / 2, 0.5);

        /// <summary>t kritis dua sisi (mis. alpha 0.05 → t tabel).</summary>
        public static double TInvTwoTailed(double alpha, double df)
        {
            double lo = 0, hi = 1000;
            for (int i = 0; i < 200; i++)
            {
                double mid = (lo + hi) / 2;
                if (TTwoTailed(mid, df) > alpha) lo = mid; else hi = mid;
            }
            return (lo + hi) / 2;
        }

        /// <summary>r tabel (nilai kritis korelasi Pearson) untuk n sampel, uji dua sisi.</summary>
        public static double RTable(int n, double alpha = 0.05)
        {
            int df = n - 2;
            if (df < 1) return double.NaN;
            double t = TInvTwoTailed(alpha, df);
            return t / Math.Sqrt(df + t * t);
        }

        /// <summary>P(F &gt; f) — nilai p uji F.</summary>
        public static double FSurvival(double f, double df1, double df2)
        {
            if (f <= 0) return 1;
            return IncompleteBeta(df2 / (df2 + df1 * f), df2 / 2, df1 / 2);
        }

        /// <summary>P(X² &gt; x) — nilai p uji chi-square.</summary>
        public static double ChiSquareSurvival(double x, double df)
        {
            if (x <= 0) return 1;
            return 1 - GammaP(df / 2, x / 2);
        }
    }
}
