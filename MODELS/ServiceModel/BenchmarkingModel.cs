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
        [MaxLength(100)]
        public string? NamaPerusahaan { get; set; }
        [MaxLength(50)]
        public string? Negara { get; set; }
        [MaxLength(100)]
        public string? Rasio { get; set; }
    }

    public class BenchmarkingTahunModel
    {
        [Key]
        public int BenchmarkingTahunId { get; set; }
        public int BenchmarkingId { get; set; }
        public int Tahun { get; set; }
        public double Rasio { get; set; }
    }

    // public class BenchmarkingResult
    // {
    //     public string? keterangan { get; set; }
    //     public double? tahun2019 { get; set; } = 0;
    //     public double? tahun2020 { get; set; } = 0;
    //     public double? tahun2021 { get; set; } = 0;
    //     public double? tahun2022 { get; set; } = 0;
    //     public double? tahun2023 { get; set; } = 0;
    //     public double? tahun2024 { get; set; } = 0;
    //     public double? rasio2025 { get; set; } = 0;
    //     public double? rasio2026 { get; set; } = 0;
    //     public double? rasio2027 { get; set; } = 0;
    //     public double? rasio2028 { get; set; } = 0;
    //     public double? rasio2029 { get; set; } = 0;
    //     public double? rasio2030 { get; set; } = 0;
    // }

}
