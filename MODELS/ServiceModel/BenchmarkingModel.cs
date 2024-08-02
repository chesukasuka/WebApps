using System.ComponentModel.DataAnnotations;

namespace WebApps.Models.ServiceModel
{
    public class BenchmarkingModel
    {
        [Key]
        public int id { get; set; }
        public string? klasifikasi_usaha { get; set; }
        public string? jenis_kegiatan_usaha { get; set; }
        public string? nama_perusahaan { get; set; }
        public string? negara { get; set; }
        public string? rasio { get; set; }

        public double tahun2019 { get; set; } = 0;
        public double tahun2020 { get; set; } = 0;
        public double tahun2021 { get; set; } = 0;
        public double tahun2022 { get; set; } = 0;
        public double tahun2023 { get; set; } = 0;
        public double tahun2024 { get; set; } = 0;
        public double rasio2025 { get; set; } = 0;
        public double rasio2026 { get; set; } = 0;
        public double rasio2027 { get; set; } = 0;
        public double rasio2028 { get; set; } = 0;
        public double rasio2029 { get; set; } = 0;
        public double rasio2030 { get; set; } = 0;
    }

    // public class BenchmarkingTahunModel
    // {
    //     [Key]
    //     public int id { get; set; }
    //     public int benchmarking_id { get; set; }
    //     public string? tahun { get; set; }
    //     public double? rasio { get; set; }
    // }

    public class BenchmarkingHitungModel
    {
        public int tahun_awal { get; set; }
        public int tahun_akhir { get; set; }
        public string rasio { get; set; }
    }

}
