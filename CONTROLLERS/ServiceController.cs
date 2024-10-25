using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json.Linq;
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
                //ViewBag.sliderValue = new int[] { UtilityController.dtYear-1, UtilityController.dtYear-1 };

                ViewBag.tahun = new string[] { "1 Tahun", "3 Tahun", "5 Tahun"};
                ViewBag.tahunpajak = new string[] { "2018", "2019", "2020", "2021", "2022", "2023", "2024" };

                ViewBag.jenis = _context.Benchmarking
                .Select(z => new BenchmarkingModel { JenisKegiatanUsaha = z.JenisKegiatanUsaha })
                .Where(z => z.JenisKegiatanUsaha != null)
                .Distinct()
                .ToList();

                ViewBag.klasifikasi = _context.Benchmarking
                .Select(z => new BenchmarkingModel { KlasifikasiUsaha = z.KlasifikasiUsaha, JenisKegiatanUsaha = z.JenisKegiatanUsaha })
                .Where(z => z.KlasifikasiUsaha != null)
                .Distinct()
                .ToList();

                ViewBag.ratio = _context.Benchmarking
                .Select(z => new BenchmarkingModel{ Rasio = z.Rasio , KlasifikasiUsaha = z.KlasifikasiUsaha})
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

        [HttpGet]
        public IActionResult Klasifikasi(string param)
        {
            var oResult = _context.Benchmarking
            .Select(z => new BenchmarkingModel { KlasifikasiUsaha = z.KlasifikasiUsaha })
            .Where(z => z.KlasifikasiUsaha != null)
            //.Distinct()
            .ToList();

            return Json(oResult);
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
                                oListData.Add(" " + (tahun1+i).ToString() , Convert.ToDouble(result.GetValue(i+2)).ToString("F2") );
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

                if (oData.Count != 0)
                {
                    var oListHeader = new List<Dictionary<string, object>>();
                    var oListData = new Dictionary<string, object>();

                    for (int j = 0; j < 5; j++)
                    {
                        var oLoop = tahun2 - tahun1 + 1;
                        oListData = new Dictionary<string, object>();
                        if (j == 0)
                        {
                            oListData.Add("Keterangan", "Minimum");
                            for (int i = 0; i < oLoop; i++)
                            {
                                var oResData = oData.Select(z => z[" " + (tahun1 + i).ToString()]).ToList();
                                oListData.Add(" " + (tahun1 + i).ToString(), GetPercentile(oResData, 0).ToString("F2"));
                            }
                        }
                        if (j == 1)
                        {
                            oListData.Add("Keterangan", "Kuartil 1");
                            for (int i = 0; i < oLoop; i++)
                            {
                                var oResData = oData.Select(z => z[" " + (tahun1 + i).ToString()]).ToList();
                                oListData.Add(" " + (tahun1 + i).ToString(), GetPercentile(oResData, 25).ToString("F2"));
                            }
                        }
                        if (j == 2)
                        {
                            oListData.Add("Keterangan", "Kuartil 2");
                            for (int i = 0; i < oLoop; i++)
                            {
                                var oResData = oData.Select(z => z[" " + (tahun1 + i).ToString()]).ToList();
                                oListData.Add(" " + (tahun1 + i).ToString(), GetPercentile(oResData, 50).ToString("F2"));
                            }
                        }
                        if (j == 3)
                        {
                            oListData.Add("Keterangan", "Kuartil 3");
                            for (int i = 0; i < oLoop; i++)
                            {
                                var oResData = oData.Select(z => z[" " + (tahun1 + i).ToString()]).ToList();
                                oListData.Add(" " + (tahun1 + i).ToString(), GetPercentile(oResData, 75).ToString("F2"));
                            }
                        }
                        if (j == 4)
                        {
                            oListData.Add("Keterangan", "Maksimum");
                            for (int i = 0; i < oLoop; i++)
                            {
                                var oResData = oData.Select(z => z[" " + (tahun1 + i).ToString()]).ToList();
                                oListData.Add(" " + (tahun1 + i).ToString(), GetPercentile(oResData, 100).ToString("F2"));
                            }
                        }
                        oListHeader.Add(oListData);
                    }
                    oResult = oListHeader;
                }
                else
                {
                    oResult = null;
                }

            }
            catch (System.Exception)
            {
                throw;
            }
            return Json(oResult);
        }

        public static double GetPercentile(List<object> dataNew, double percentile)
        {
            
            var sortedValues = new List<double>();
            foreach (object oData in dataNew)
            {
                sortedValues.Add(Convert.ToDouble(oData));
            }

            // Sort the list
            sortedValues.Sort();

            // Calculate the index
            int N = sortedValues.Count;
            double rank = (percentile / 100) * (N - 1);
            int lowerIndex = (int)Math.Floor(rank);
            int upperIndex = (int)Math.Ceiling(rank);

            // If exact index, return the value
            if (lowerIndex == upperIndex)
                return sortedValues[lowerIndex];

            // Interpolate between the two surrounding values
            double fraction = rank - lowerIndex;
            return sortedValues[lowerIndex] + fraction * (sortedValues[upperIndex] - sortedValues[lowerIndex]);
        }



    }
}
