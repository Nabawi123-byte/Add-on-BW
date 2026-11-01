using System;
using System.Collections.Generic;
using System.Linq;

namespace MasterBW.Core.Finance
{
    public sealed class ScheduleRow
    {
        public int Period { get; set; }
        public DateTime? Date { get; set; }
        public double Opening { get; set; }
        public double Payment { get; set; }
        public double Interest { get; set; }
        public double Principal { get; set; }
        public double Closing { get; set; }
    }

    public sealed class DepreciationRow
    {
        public int Year { get; set; }
        public double Opening { get; set; }
        public double Expense { get; set; }
        public double Accumulated { get; set; }
        public double BookValue { get; set; }
    }

    /// <summary>Perhitungan keuangan: pinjaman, penyusutan, investasi, BEP, rasio.</summary>
    public static class FinanceCalc
    {
        /// <summary>Angsuran anuitas per periode (setara PMT Excel, nilai positif).</summary>
        public static double Pmt(double ratePerPeriod, int periods, double principal)
        {
            if (periods <= 0) throw new ArgumentException("Jumlah periode harus > 0.");
            if (Math.Abs(ratePerPeriod) < 1e-12) return principal / periods;
            return principal * ratePerPeriod / (1 - Math.Pow(1 + ratePerPeriod, -periods));
        }

        /// <summary>Jadwal angsuran. method: annuity | flat | effective (pokok tetap, bunga menurun).</summary>
        public static List<ScheduleRow> LoanSchedule(double principal, double annualRatePercent, int months, string method = "annuity", DateTime? start = null)
        {
            var rows = new List<ScheduleRow>();
            double r = annualRatePercent / 100.0 / 12.0;
            double balance = principal;
            double annuity = Pmt(r, months, principal);
            double flatInterest = principal * r;
            double fixedPrincipal = principal / months;
            for (int i = 1; i <= months; i++)
            {
                var row = new ScheduleRow { Period = i, Opening = balance, Date = start?.AddMonths(i - 1) };
                switch ((method ?? "annuity").ToLowerInvariant())
                {
                    case "flat":
                        row.Interest = flatInterest;
                        row.Principal = fixedPrincipal;
                        break;
                    case "effective":
                        row.Interest = balance * r;
                        row.Principal = fixedPrincipal;
                        break;
                    default:
                        row.Interest = balance * r;
                        row.Principal = annuity - row.Interest;
                        break;
                }
                if (i == months) row.Principal = balance; // bereskan pembulatan
                row.Payment = row.Principal + row.Interest;
                balance -= row.Principal;
                row.Closing = Math.Abs(balance) < 1e-6 ? 0 : balance;
                rows.Add(row);
            }
            return rows;
        }

        /// <summary>method: straight_line | declining_balance | double_declining | sum_of_years.</summary>
        public static List<DepreciationRow> Depreciation(double cost, double salvage, int lifeYears, string method = "straight_line")
        {
            if (lifeYears <= 0) throw new ArgumentException("Umur ekonomis harus > 0.");
            var rows = new List<DepreciationRow>();
            double book = cost, acc = 0;
            int soy = lifeYears * (lifeYears + 1) / 2;
            double dbRate = salvage > 0 && cost > 0 ? 1 - Math.Pow(salvage / cost, 1.0 / lifeYears) : 2.0 / lifeYears;
            for (int y = 1; y <= lifeYears; y++)
            {
                double exp;
                switch ((method ?? "").ToLowerInvariant())
                {
                    case "declining_balance": exp = book * dbRate; break;
                    case "double_declining": exp = book * (2.0 / lifeYears); break;
                    case "sum_of_years": exp = (cost - salvage) * (lifeYears - y + 1) / soy; break;
                    default: exp = (cost - salvage) / lifeYears; break;
                }
                if (book - exp < salvage || y == lifeYears) exp = book - salvage;
                if (exp < 0) exp = 0;
                var row = new DepreciationRow { Year = y, Opening = book, Expense = exp };
                acc += exp;
                book -= exp;
                row.Accumulated = acc;
                row.BookValue = book;
                rows.Add(row);
            }
            return rows;
        }

        /// <summary>NPV dengan arus kas periode 0 tidak didiskonto (berbeda dengan NPV Excel yang mulai periode 1).</summary>
        public static double Npv(double ratePercent, IList<double> cashflows)
        {
            double r = ratePercent / 100.0, npv = 0;
            for (int t = 0; t < cashflows.Count; t++) npv += cashflows[t] / Math.Pow(1 + r, t);
            return npv;
        }

        /// <summary>IRR (persen) dengan Newton-Raphson + fallback bisection. null bila tidak ada solusi.</summary>
        public static double? Irr(IList<double> cashflows)
        {
            if (cashflows.Count < 2 || !cashflows.Any(c => c < 0) || !cashflows.Any(c => c > 0)) return null;
            double f(double r) { double s = 0; for (int t = 0; t < cashflows.Count; t++) s += cashflows[t] / Math.Pow(1 + r, t); return s; }
            double df(double r) { double s = 0; for (int t = 1; t < cashflows.Count; t++) s -= t * cashflows[t] / Math.Pow(1 + r, t + 1); return s; }
            double x = 0.1;
            for (int i = 0; i < 100; i++)
            {
                double fx = f(x), d = df(x);
                if (Math.Abs(d) < 1e-14) break;
                double nx = x - fx / d;
                if (nx <= -0.9999) nx = (x - 0.9999) / 2;
                if (Math.Abs(nx - x) < 1e-12) return nx * 100;
                x = nx;
            }
            double lo = -0.9999, hi = 10;
            double flo = f(lo), fhi = f(hi);
            if (flo * fhi > 0) return null;
            for (int i = 0; i < 300; i++)
            {
                double mid = (lo + hi) / 2, fm = f(mid);
                if (Math.Abs(fm) < 1e-9) return mid * 100;
                if (flo * fm < 0) { hi = mid; fhi = fm; } else { lo = mid; flo = fm; }
            }
            return (lo + hi) / 2 * 100;
        }

        /// <summary>Payback period (tahun, dengan interpolasi). null jika tidak pernah balik modal.</summary>
        public static double? PaybackPeriod(IList<double> cashflows, double? discountRatePercent = null)
        {
            double cum = 0;
            for (int t = 0; t < cashflows.Count; t++)
            {
                double cf = discountRatePercent.HasValue ? cashflows[t] / Math.Pow(1 + discountRatePercent.Value / 100, t) : cashflows[t];
                double prev = cum;
                cum += cf;
                if (t > 0 && prev < 0 && cum >= 0) return (t - 1) + (-prev / cf);
            }
            return null;
        }

        public static double ProfitabilityIndex(double ratePercent, IList<double> cashflows)
        {
            double initial = -cashflows[0];
            if (initial <= 0) return double.NaN;
            return (Npv(ratePercent, cashflows) + initial) / initial;
        }

        public static (double units, double rupiah) BreakEven(double fixedCost, double price, double variableCost)
        {
            var margin = price - variableCost;
            if (margin <= 0) throw new ArgumentException("Harga jual harus lebih besar dari biaya variabel per unit.");
            var units = fixedCost / margin;
            return (units, units * price);
        }

        /// <summary>Rasio keuangan standar. Kunci input: aset_lancar, utang_lancar, persediaan, kas, total_aset, total_utang, ekuitas, penjualan, laba_kotor, laba_operasi, laba_bersih.</summary>
        public static List<(string group, string name, double? value, string unit, string formula)> Ratios(IDictionary<string, double> v)
        {
            double? G(string k) => v.TryGetValue(k, out var x) ? x : (double?)null;
            double? Div(double? a, double? b) => a.HasValue && b.HasValue && Math.Abs(b.Value) > 1e-12 ? a / b : null;
            double? Pct(double? x) => x * 100;
            return new List<(string, string, double?, string, string)>
            {
                ("Likuiditas", "Current Ratio", Div(G("aset_lancar"), G("utang_lancar")), "x", "Aset lancar / Utang lancar"),
                ("Likuiditas", "Quick Ratio", Div(G("aset_lancar") - (G("persediaan") ?? 0), G("utang_lancar")), "x", "(Aset lancar − Persediaan) / Utang lancar"),
                ("Likuiditas", "Cash Ratio", Div(G("kas"), G("utang_lancar")), "x", "Kas / Utang lancar"),
                ("Solvabilitas", "Debt to Asset Ratio", Pct(Div(G("total_utang"), G("total_aset"))), "%", "Total utang / Total aset"),
                ("Solvabilitas", "Debt to Equity Ratio", Div(G("total_utang"), G("ekuitas")), "x", "Total utang / Ekuitas"),
                ("Profitabilitas", "Gross Profit Margin", Pct(Div(G("laba_kotor"), G("penjualan"))), "%", "Laba kotor / Penjualan"),
                ("Profitabilitas", "Operating Margin", Pct(Div(G("laba_operasi"), G("penjualan"))), "%", "Laba operasi / Penjualan"),
                ("Profitabilitas", "Net Profit Margin", Pct(Div(G("laba_bersih"), G("penjualan"))), "%", "Laba bersih / Penjualan"),
                ("Profitabilitas", "Return on Assets (ROA)", Pct(Div(G("laba_bersih"), G("total_aset"))), "%", "Laba bersih / Total aset"),
                ("Profitabilitas", "Return on Equity (ROE)", Pct(Div(G("laba_bersih"), G("ekuitas"))), "%", "Laba bersih / Ekuitas"),
                ("Aktivitas", "Total Asset Turnover", Div(G("penjualan"), G("total_aset")), "x", "Penjualan / Total aset"),
            };
        }
    }
}
