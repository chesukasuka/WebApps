using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Syncfusion.EJ2.Linq;
using System.Collections;
using System.Diagnostics;
using WebApps.Models;
using WebApps.Models.ServiceModel;

namespace WebApps.Controllers
{
    public class ServiceController : Controller
    {
        private readonly ILogger<ServiceController> _logger;
        private readonly MasterDbContext _context;

        public ServiceController(ILogger<ServiceController> logger, MasterDbContext context)
        {
            _logger = logger;
            _context = context;
        }

        public IActionResult Benchmarking()
        {
            var msg = "";
            try
            {
                ViewBag.sliderValue = new int[] { UtilityController.dtYear-1, UtilityController.dtYear-1 };

                ViewBag.jenis = _context.Benchmarking
                .Select(z => new BenchmarkingModel{ JenisKegiatanUsaha = z.JenisKegiatanUsaha })
                .Where(z => z.JenisKegiatanUsaha != null)
                .Distinct()
                .ToList();                

                ViewBag.klasifikasi = _context.Benchmarking
                .Select(z => new BenchmarkingModel{ KlasifikasiUsaha = z.KlasifikasiUsaha })
                .Where(z => z.KlasifikasiUsaha != null)
                .Distinct()
                .ToList();                

                ViewBag.ratio = _context.Benchmarking
                .Select(z => new BenchmarkingModel{ Rasio = z.Rasio })
                .Where(z => z.Rasio != null)
                .Distinct()
                .ToList();                

                // ViewBag.dataSource = _context.Benchmarking
                // .ToList();
                ViewBag.dataSource = null;

            }
            catch (System.Exception e)
            {
                msg = e.Message;                
            }

            return View();
        }

        private List<Dictionary<string, object>> Searchbenchmark(string rasio, string jenis, string klasifikasi, int tahun1, int tahun2){
            var oResult = new List<Dictionary<string, object>>();
            try
            {   
                var oLoop = tahun2 - tahun1 + 1;
                var oList = new List<Dictionary<string, object>>();

                var sTahun = "[" + tahun1.ToString() + "]";
                for (int i = tahun1; i < tahun2; i++){
                    sTahun = sTahun + ",[" + (i+1).ToString() + "]";
                }
                var oListHeader = new List<Dictionary<string, object>>();
                                                

                var sql = "";
                sql = ""
                    + " select a.NamaPerusahaan, a.Negara, " + sTahun + "              "
                    + " from benchmarking a                                 "
                    + " left join                                           "
                    + " (                                                   "
                    + " SELECT                                              "
                    + " *                                                   "
                    + " FROM                                                "
                    + "     (SELECT BenchmarkingId, Tahun, Rasio            "
                    + "     FROM BenchmarkingTahun) as SourceTable          "
                    + " PIVOT                                               "
                    + " (                                                   "
                    + "     MAX(Rasio)                                      "
                    + "     FOR Tahun IN (" + sTahun + ")                   "
                    + " ) as PivotTable                                     "
                    + " ) b                                                 "
                    + " on a.BenchmarkingId=b.BenchmarkingId                "
                    + " WHERE 1 = 1                                         ";

                if(rasio != null){
                    sql = sql + " AND Rasio = '" + rasio + "' ";
                }
                if(jenis != null){
                    sql = sql + " AND JenisKegiatanUsaha = '" + jenis + "' ";
                }
                if(jenis != null){
                    sql = sql + " AND KlasifikasiUsaha = '" + klasifikasi + "' ";
                }

                using (var command = _context.Database.GetDbConnection().CreateCommand())
                {
                    command.CommandText = sql;
                    _context.Database.OpenConnection();

                    using (var result = command.ExecuteReader())
                    {
                        while (result.Read())
                        {
                            bool bInsert = true;
                            var oListData = new Dictionary<string, object>();
                            oListData.Add("Nama Perusahaan", result.GetValue(0));
                            oListData.Add("Negara", result.GetValue(1));
                            for (int i = 0; i < oLoop; i++){
                                oListData.Add(" " + (tahun1+i).ToString() ,result.GetValue(i+2));
                                if(result.GetValue(i+2).ToString() == "0"){
                                    bInsert = false;
                                    continue;
                                }
                            }
                            if(bInsert){
                                oListHeader.Add(oListData);
                            }
                        }
                    }
                }

                oResult = oListHeader;
            }
            catch (System.Exception)
            {                
                throw;
            }

            return oResult;
        }
        public IActionResult Hitung(string rasio, string jenis, string klasifikasi, int tahun1, int tahun2)
        {
            var data = ViewBag.sliderValue;
            dynamic oResult;
            try
            {   
                oResult = Searchbenchmark(rasio,jenis,klasifikasi,tahun1,tahun2);
            }
            catch (System.Exception)
            {                
                throw;
            }
            return Json(oResult);
        }

        public IActionResult Hitung2(string rasio, string jenis, string klasifikasi, int tahun1, int tahun2)
        {
            dynamic oResult;
            try
            {
                var oData = Searchbenchmark(rasio,jenis,klasifikasi,tahun1,tahun2);
                
                var oListHeader = new List<Dictionary<string, object>>();
                var oListData = new Dictionary<string, object>();
                
                for (int j = 0; j < 5; j++){
                    var oLoop = tahun2 - tahun1 + 1;
                    oListData = new Dictionary<string, object>();
                    if(j == 0){
                        oListData.Add("Keterangan","Minimal");
                        for (int i = 0; i < oLoop; i++){
                            var oResData = oData.OrderBy(dict => dict[" " + (tahun1 + i).ToString()]).First();
                            oListData.Add(" " + (tahun1 + i).ToString(), oResData[" " + (tahun1 + i).ToString()]);
                        }
                    }
                    if(j == 1){
                        oListData.Add("Keterangan","Kuartil 1");
                        for (int i = 0; i < oLoop; i++){
                            var oResData = oData.Select(z => z[" " + (tahun1 + i).ToString()] ).ToList();                            
                            oListData.Add(" " + (tahun1 + i).ToString(), GetPercentile(oResData, 0.25).ToString("F2").Replace(".00","") );
                        }
                    }
                    if(j == 2){
                        oListData.Add("Keterangan","Kuartil 2");
                        for (int i = 0; i < oLoop; i++){
                            var oResData = oData.Select(z => z[" " + (tahun1 + i).ToString()] ).ToList();                            
                            oListData.Add(" " + (tahun1 + i).ToString(), GetPercentile(oResData, 0.50).ToString("F2").Replace(".00","") );
                        }                        
                    }
                    if(j == 3){
                        oListData.Add("Keterangan","Kuartil 3");
                        for (int i = 0; i < oLoop; i++){
                            var oResData = oData.Select(z => z[" " + (tahun1 + i).ToString()] ).ToList();                            
                            oListData.Add(" " + (tahun1 + i).ToString(), GetPercentile(oResData, 0.75).ToString("F2").Replace(".00","") );
                        }
                    }
                    if(j == 4){
                        oListData.Add("Keterangan","Maksimum");                        
                        for (int i = 0; i < oLoop; i++){
                            var oResData = oData.OrderBy(dict => dict[" " + (tahun1 + i).ToString()]).Last();
                            oListData.Add(" " + (tahun1 + i).ToString(), oResData[" " + (tahun1 + i).ToString()]);
                        }
                    }
                    oListHeader.Add(oListData);
                }

                oResult = oListHeader;


                // var oModel = new BenchmarkingResult();
                // oModel.keterangan = "Minimal";
                // oModel.tahun2019 = _context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Min(z => z.tahun2019);
                // oModel.tahun2020 = _context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Min(z => z.tahun2020);
                // oModel.tahun2021 = _context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Min(z => z.tahun2021);
                // oModel.tahun2022 = _context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Min(z => z.tahun2022);
                // oModel.tahun2023 = _context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Min(z => z.tahun2023);
                // oModel.tahun2024 = _context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Min(z => z.tahun2024);
                // oResult.Add(oModel);

                // oModel = new BenchmarkingResult();
                // oModel.keterangan = "Kuartil 1";
                // oModel.tahun2019 = CalQ1(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2019).ToList());
                // oModel.tahun2020 = CalQ1(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2020).ToList());
                // oModel.tahun2021 = CalQ1(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2021).ToList());
                // oModel.tahun2022 = CalQ1(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2022).ToList());
                // oModel.tahun2023 = CalQ1(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2023).ToList());
                // oModel.tahun2024 = CalQ1(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2024).ToList());
                // oResult.Add(oModel);

                // oModel = new BenchmarkingResult();
                // oModel.keterangan = "Kuartil 2";
                // oModel.tahun2019 = CalQ2(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2019).ToList());
                // oModel.tahun2020 = CalQ2(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2020).ToList());
                // oModel.tahun2021 = CalQ2(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2021).ToList());
                // oModel.tahun2022 = CalQ2(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2022).ToList());
                // oModel.tahun2023 = CalQ2(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2023).ToList());
                // oModel.tahun2024 = CalQ2(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2024).ToList());
                // oResult.Add(oModel);

                // oModel = new BenchmarkingResult();
                // oModel.keterangan = "Kuartil 3";
                // oModel.tahun2019 = CalQ3(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2019).ToList());
                // oModel.tahun2020 = CalQ3(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2020).ToList());
                // oModel.tahun2021 = CalQ3(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2021).ToList());
                // oModel.tahun2022 = CalQ3(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2022).ToList());
                // oModel.tahun2023 = CalQ3(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2023).ToList());
                // oModel.tahun2024 = CalQ3(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2024).ToList());
                // oResult.Add(oModel);

                // oModel = new BenchmarkingResult();
                // oModel.keterangan = "Maksimum";
                // oModel.tahun2019 = _context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Max(z => z.tahun2019);
                // oModel.tahun2020 = _context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Max(z => z.tahun2020);
                // oModel.tahun2021 = _context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Max(z => z.tahun2021);
                // oModel.tahun2022 = _context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Max(z => z.tahun2022);
                // oModel.tahun2023 = _context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Max(z => z.tahun2023);
                // oModel.tahun2024 = _context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Max(z => z.tahun2024);
                // oResult.Add(oModel);

            }
            catch (System.Exception)
            {
                throw;
            }
            return Json(oResult);
        }

        static double GetPercentile(List<object> dataNew, double percentile)
        {
            double result = 0;
            var data = new List<double>();
            foreach (object oData in dataNew){
                data.Add(Convert.ToDouble(oData));
            }
            int n = data.Count;
            double rank = (percentile * (n - 1)) + 1;
            int intPart = (int)rank;
            double fracPart = rank - intPart;

            if (intPart >= n)
                return data[n - 1];
            if (intPart == 0)
                return data[0];

            result = data[intPart - 1] + fracPart * (data[intPart] - data[intPart - 1]);
                
            return result;
        }


    }
}
