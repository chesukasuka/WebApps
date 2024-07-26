namespace WebApps.Models.ServiceModel
{
    public class BenchmarkingModel
    {
        public int id { get; set; }
        public string klasifikasi_usaha { get; set; }
        public string jenis_kegiatan_usaha { get; set; }
        public string nama_perusahaan { get; set; }
        public string negara { get; set; }
        public string rasio { get; set; }
    }

    public class BenchmarkingTahunModel
    {
        public int id { get; set; }
        public int benchmarking_id { get; set; }
        public string tahun { get; set; }
        public decimal rasio { get; set; }
    }

}
