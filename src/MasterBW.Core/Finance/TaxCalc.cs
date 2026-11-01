using System;
using System.Collections.Generic;

namespace MasterBW.Core.Finance
{
    public sealed class TaxLine
    {
        public string Label { get; set; }
        public double Amount { get; set; }
        public string Note { get; set; }
        public TaxLine(string label, double amount, string note = "") { Label = label; Amount = amount; Note = note; }
    }

    /// <summary>
    /// Kalkulator pajak Indonesia (estimasi). Tarif mengacu UU HPP No. 7/2021.
    /// PENTING: untuk pemotongan bulanan resmi sejak 2024 gunakan TER (PP 58/2023); perhitungan tahunan di sini
    /// memakai tarif Pasal 17 dan cocok untuk perhitungan setahun/masa Desember. Selalu verifikasi dengan aturan terbaru.
    /// </summary>
    public static class TaxCalc
    {
        public const double PtkpWp = 54_000_000;
        public const double PtkpKawin = 4_500_000;
        public const double PtkpTanggungan = 4_500_000;

        /// <summary>PTKP setahun dari status "TK/0", "K/2", dst.</summary>
        public static double Ptkp(string status)
        {
            status = (status ?? "TK/0").ToUpperInvariant().Replace(" ", "");
            var parts = status.Split('/');
            bool kawin = parts[0] == "K";
            int tanggungan = parts.Length > 1 && int.TryParse(parts[1], out var t) ? Math.Max(0, Math.Min(3, t)) : 0;
            return PtkpWp + (kawin ? PtkpKawin : 0) + tanggungan * PtkpTanggungan;
        }

        /// <summary>Tarif progresif Pasal 17 (UU HPP) atas PKP setahun.</summary>
        public static double Pasal17(double pkp)
        {
            if (pkp <= 0) return 0;
            var brackets = new[] { (60_000_000d, 0.05), (250_000_000d, 0.15), (500_000_000d, 0.25), (5_000_000_000d, 0.30), (double.MaxValue, 0.35) };
            double tax = 0, lower = 0;
            foreach (var (upper, rate) in brackets)
            {
                if (pkp <= lower) break;
                tax += (Math.Min(pkp, upper) - lower) * rate;
                lower = upper;
            }
            return tax;
        }

        /// <summary>PPh 21 pegawai tetap setahun dari bruto bulanan.</summary>
        public static List<TaxLine> Pph21Annual(double grossMonthly, string ptkpStatus, bool hasNpwp = true, double iuranPensiunMonthly = 0)
        {
            var lines = new List<TaxLine>();
            double brutoYear = grossMonthly * 12;
            double biayaJabatan = Math.Min(brutoYear * 0.05, 6_000_000);
            double iuran = iuranPensiunMonthly * 12;
            double neto = brutoYear - biayaJabatan - iuran;
            double ptkp = Ptkp(ptkpStatus);
            double pkp = Math.Max(0, Math.Floor((neto - ptkp) / 1000) * 1000);
            double tax = Pasal17(pkp);
            double surcharge = hasNpwp ? 0 : tax * 0.2;
            lines.Add(new TaxLine("Penghasilan bruto setahun", brutoYear, "bruto bulanan × 12"));
            lines.Add(new TaxLine("Biaya jabatan", biayaJabatan, "5%, maks. Rp6.000.000/tahun"));
            if (iuran > 0) lines.Add(new TaxLine("Iuran pensiun/JHT", iuran));
            lines.Add(new TaxLine("Penghasilan neto setahun", neto));
            lines.Add(new TaxLine("PTKP (" + ptkpStatus + ")", ptkp));
            lines.Add(new TaxLine("PKP (dibulatkan ribuan ke bawah)", pkp));
            lines.Add(new TaxLine("PPh 21 setahun (tarif Pasal 17)", tax, "5%/15%/25%/30%/35%"));
            if (!hasNpwp) lines.Add(new TaxLine("Tambahan 20% (tanpa NPWP)", surcharge));
            lines.Add(new TaxLine("Total PPh 21 setahun", tax + surcharge));
            lines.Add(new TaxLine("Rata-rata PPh 21 per bulan", (tax + surcharge) / 12, "estimasi; pemotongan Jan–Nov memakai TER"));
            lines.Add(new TaxLine("Gaji bersih per bulan (estimasi)", grossMonthly - (tax + surcharge) / 12 - iuranPensiunMonthly));
            return lines;
        }

        /// <summary>
        /// PPN. Sejak 2025 tarif 12%; untuk barang/jasa non-mewah DPP = 11/12 × harga, sehingga efektif 11%.
        /// </summary>
        public static List<TaxLine> Ppn(double hargaJual, bool barangMewah = false)
        {
            double dpp = barangMewah ? hargaJual : hargaJual * 11.0 / 12.0;
            double ppn = dpp * 0.12;
            return new List<TaxLine>
            {
                new TaxLine("Harga jual / penggantian", hargaJual),
                new TaxLine("DPP" + (barangMewah ? "" : " nilai lain (11/12)"), dpp),
                new TaxLine("PPN 12%", ppn, barangMewah ? "barang mewah" : "efektif 11% dari harga jual"),
                new TaxLine("Total tagihan", hargaJual + ppn),
            };
        }

        /// <summary>PPh 23: 2% jasa / 15% dividen-bunga-royalti; tanpa NPWP tarif 100% lebih tinggi.</summary>
        public static List<TaxLine> Pph23(double bruto, bool jasa = true, bool hasNpwp = true)
        {
            double rate = (jasa ? 0.02 : 0.15) * (hasNpwp ? 1 : 2);
            return new List<TaxLine>
            {
                new TaxLine("Jumlah bruto", bruto),
                new TaxLine("Tarif PPh 23", rate * 100, "%"),
                new TaxLine("PPh 23 dipotong", bruto * rate),
                new TaxLine("Diterima bersih", bruto - bruto * rate),
            };
        }

        /// <summary>PPh Final UMKM 0,5% (PP 55/2022); WP orang pribadi bebas atas omzet s.d. Rp500 juta setahun.</summary>
        public static List<TaxLine> PphFinalUmkm(double omzetBulanan, double omzetKumulatifSebelumnya = 0, bool orangPribadi = true)
        {
            double bebas = orangPribadi ? 500_000_000 : 0;
            double sisaBebas = Math.Max(0, bebas - omzetKumulatifSebelumnya);
            double kenaPajak = Math.Max(0, omzetBulanan - sisaBebas);
            return new List<TaxLine>
            {
                new TaxLine("Omzet bulan ini", omzetBulanan),
                new TaxLine("Sisa omzet tidak kena pajak", Math.Min(sisaBebas, omzetBulanan), orangPribadi ? "Rp500 juta/tahun untuk WP OP" : ""),
                new TaxLine("Omzet kena pajak", kenaPajak),
                new TaxLine("PPh Final 0,5%", kenaPajak * 0.005),
            };
        }
    }
}
