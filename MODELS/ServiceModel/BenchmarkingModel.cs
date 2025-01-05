using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace WebApps.Models.ServiceModel
{
    public class BenchmarkingModel
    {
        [Key]
        public int BenchmarkingId { get; set; }
        [MaxLength(100)]
        public string? KlasifikasiUsaha { get; set; }
        [MaxLength(100)]
        public string? JenisKegiatanUsaha { get; set; }
        [MaxLength(150)]
        public string? NamaPerusahaan { get; set; }
        [MaxLength(50)]
        public string? Negara { get; set; }
        [MaxLength(100)]
        public string? Rasio { get; set; }
        [MaxLength(100)]
        public string? SubKlasifikasiUsaha { get; set; }
        [MaxLength(100)]
        public string? Metode { get; set; }
    }

    public class BenchmarkingTahunModel
    {
        [Key]
        public int BenchmarkingTahunId { get; set; }
        public int BenchmarkingId { get; set; }
        public int Tahun { get; set; }
        public double Rasio { get; set; }
    }

    //public class BenchmarkingResult
    //{
    //    public string? keterangan { get; set; }
    //    public double? tahun2019 { get; set; } = 0;
    //    public double? tahun2020 { get; set; } = 0;
    //    public double? tahun2021 { get; set; } = 0;
    //    public double? tahun2022 { get; set; } = 0;
    //    public double? tahun2023 { get; set; } = 0;
    //    public double? tahun2024 { get; set; } = 0;
    //    public double? rasio2025 { get; set; } = 0;
    //    public double? rasio2026 { get; set; } = 0;
    //    public double? rasio2027 { get; set; } = 0;
    //    public double? rasio2028 { get; set; } = 0;
    //    public double? rasio2029 { get; set; } = 0;
    //    public double? rasio2030 { get; set; } = 0;
    //}

    public class FakturKeluaranDaftarModel
    {
        [Key]
        public long? FakturKeluaranDaftarId { get; set; }
        public string? NamaFile { get; set; }
        public int? Jumlah { get; set; }
    }

    public class FakturKeluaranHeaderModel
    {
        [Key]
        public long? FakturKeluaranHeaderId { get; set; }
        public long? FakturKeluaranDaftarId { get; set; }
        public string? Tin { get; set; }
        public string? TaxInvoiceDate { get; set; }
        public string? TaxInvoiceOpt { get; set; }
        public string? TrxCode { get; set; }
        public string? AddInfo { get; set; }
        public string? CustomDoc { get; set; }
        public string? RefDesc { get; set; }
        public string? FacilityStamp { get; set; }
        public string? SellerIDTKU { get; set; }
        public string? BuyerTin { get; set; }
        public string? BuyerDocument { get; set; }
        public string? BuyerCountry { get; set; }
        public string? BuyerDocumentNumber { get; set; }
        public string? BuyerName { get; set; }
        public string? BuyerAdress { get; set; }
        public string? BuyerEmail { get; set; }
        public string? BuyerIDTKU { get; set; }
    }

    public class FakturKeluaranItemModel
    {
        [Key]
        public long? FakturKeluaranItemId { get; set; }
        public long? FakturKeluaranHeaderId { get; set; }
        public string? Opt { get; set; }
        public string? Code { get; set; }
        public string? Name { get; set; }
        public string? Unit { get; set; }
        public double? Price { get; set; }
        public double? Qty { get; set; }
        public double? TotalDiscount { get; set; }
        public double? TaxBase { get; set; }
        public double? OtherTaxBase { get; set; }
        public string? VATRate { get; set; }
        public double? VAT { get; set; }
        public string? STLGRate { get; set; }
        public double? STLG { get; set; }
    }

}
