using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Syncfusion.EJ2.Linq;
using Syncfusion.XlsIO;
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

            oResDetail = oResult;
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

        public IActionResult Hitung2(string rasio, string jenis, string klasifikasi, int tahun1, int tahun2
            , string penjualan
            , string hargapokokpenjualan
            , string labakotor
            , string bebanoperasional
            , string labaoperasional
            , string testedparty
            , string tahun
            )
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
                                var oResData = oData.OrderBy(dict => dict[" " + (tahun1 + i).ToString()]).First();
                                oListData.Add(" " + (tahun1 + i).ToString(), oResData[" " + (tahun1 + i).ToString()]);
                            }
                        }
                        if (j == 1)
                        {
                            oListData.Add("Keterangan", "Kuartil 1");
                            for (int i = 0; i < oLoop; i++)
                            {
                                var oResData = oData.Select(z => z[" " + (tahun1 + i).ToString()]).ToList();
                                oListData.Add(" " + (tahun1 + i).ToString(), GetPercentile(oResData, 0.25).ToString("F2").Replace(".00", ""));
                            }
                        }
                        if (j == 2)
                        {
                            oListData.Add("Keterangan", "Kuartil 2");
                            for (int i = 0; i < oLoop; i++)
                            {
                                var oResData = oData.Select(z => z[" " + (tahun1 + i).ToString()]).ToList();
                                oListData.Add(" " + (tahun1 + i).ToString(), GetPercentile(oResData, 0.50).ToString("F2").Replace(".00", ""));
                            }
                        }
                        if (j == 3)
                        {
                            oListData.Add("Keterangan", "Kuartil 3");
                            for (int i = 0; i < oLoop; i++)
                            {
                                var oResData = oData.Select(z => z[" " + (tahun1 + i).ToString()]).ToList();
                                oListData.Add(" " + (tahun1 + i).ToString(), GetPercentile(oResData, 0.75).ToString("F2").Replace(".00", ""));
                            }
                        }
                        if (j == 4)
                        {
                            oListData.Add("Keterangan", "Maksimum");
                            for (int i = 0; i < oLoop; i++)
                            {
                                var oResData = oData.OrderBy(dict => dict[" " + (tahun1 + i).ToString()]).Last();
                                oListData.Add(" " + (tahun1 + i).ToString(), oResData[" " + (tahun1 + i).ToString()]);
                            }
                        }
                        oListHeader.Add(oListData);
                    }
                    oResult = oListHeader;

                    oResSum = oListHeader;
                    sRasio = rasio;
                    sJenis = jenis;
                    sKlasifikasi = klasifikasi;
                    sTahun = (Convert.ToInt32(tahun2) + 1).ToString();
                    sPenjualan = penjualan;
                    sHargapokokpenjualan = hargapokokpenjualan;
                    sLabakotor = labakotor;
                    sBebanoperasional = bebanoperasional;
                    sLabaoperasional = labaoperasional;
                    sTestedparty = testedparty;
                    sRentang = tahun;
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

        private static List<Dictionary<string, object>>? oResDetail;
        private static List<Dictionary<string, object>>? oResSum;
        private static string sRasio = "";
        private static string sJenis = "";
        private static string sKlasifikasi = "";
        private static string sTahun = "";
        private static string sPenjualan = "";
        private static string sHargapokokpenjualan = "";
        private static string sLabakotor = "";
        private static string sBebanoperasional = "";
        private static string sLabaoperasional = "";
        private static string sTestedparty = "";
        private static string sRentang = "";

        public IActionResult ExportExcel()
        {
            var iRentang = 1;
            var excel = "";
            if (sRentang == "1 Tahun")
            {
                excel = "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";
                iRentang = 1;
            }
            else if (sRentang == "3 Tahun")
            {
                excel = "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";
                iRentang = 3;
            }
            else if (sRentang == "5 Tahun")
            {
                excel = "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";
                iRentang = 5;
            }

            byte[] excelBytes = Convert.FromBase64String(excel);
            using (MemoryStream stream = new MemoryStream(excelBytes))
            {
                // Create a new Excel workbook
                using (ExcelEngine excelEngine = new ExcelEngine())
                {
                    IApplication application = excelEngine.Excel;
                    application.DefaultVersion = ExcelVersion.Excel2016;

                    // Create a workbook with one worksheet
                    IWorkbook workbook = application.Workbooks.Open(stream);
                    IWorksheet worksheet = workbook.Worksheets[0];

                    
                    worksheet.Range["B2"].Text = ("BENCHMARKING \n" + "DATA PEMBANDING " + " " + sJenis + " " + sKlasifikasi).ToUpper();
                    worksheet.Range["E9"].Text = sJenis;
                    worksheet.Range["E12"].Text = sKlasifikasi;
                    worksheet.Range["E15"].Text = sTahun;
                    worksheet.Range["E17"].Text = sRasio;

                    worksheet.Range["H12"].Text = sPenjualan;
                    worksheet.Range["H13"].Text = sHargapokokpenjualan;
                    worksheet.Range["H14"].Text = sLabakotor;
                    worksheet.Range["H15"].Text = sBebanoperasional;
                    worksheet.Range["H16"].Text = sLabaoperasional;
                    worksheet.Range["H17"].Text = sTestedparty;

                    var header = 22;
                    var row = 23;
                    var num = 1;
                    foreach (var item in oResDetail)
                    {
                        if (oResDetail.Count > 1 && num != oResDetail.Count)
                        {
                            worksheet.InsertRow(row, 1, ExcelInsertOptions.FormatAsAfter);
                            worksheet.Range["C" + row.ToString() + ":F" + row.ToString()].Merge();
                            
                        }


                        worksheet.Range["B" + row.ToString()].Text = num.ToString();
                        worksheet.Range["C" + row.ToString()].Text = item["Nama Perusahaan"].ToString();
                        worksheet.Range["C" + row.ToString()].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;

                        worksheet.Range["G" + row.ToString()].Text = item["Negara"].ToString();

                        char letter = 'H';
                        var iTahun = Convert.ToInt32(sTahun) - iRentang;
                        for (int i = 0; i < iRentang; i++) 
                        {
                            worksheet.Range[letter + row.ToString()].Text = item[" " + iTahun.ToString()].ToString();
                            worksheet.Range[letter + row.ToString()].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;

                            worksheet.Range[letter + header.ToString()].Text = iTahun.ToString();
                            iTahun++;
                            letter++;
                        }


                        row++;
                        num++;
                    }

                    var rowSum = 27 + num - 2;
                    foreach (var item in oResSum)
                    {
                        char letter = 'H';
                        var iTahun = Convert.ToInt32(sTahun) - iRentang;
                        for (int i = 0; i < iRentang; i++)
                        {
                            worksheet.Range[letter + rowSum.ToString()].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;
                            worksheet.Range[letter + rowSum.ToString()].Text = item[" " + iTahun.ToString()].ToString();
                            iTahun++;
                            letter++;
                        }
                        rowSum++;
                    }

                    // Save the workbook to a memory stream
                    using (MemoryStream streamFile = new MemoryStream())
                    {
                        workbook.SaveAs(streamFile);
                        streamFile.Position = 0;

                        // Return the Excel file as a download
                        return File(streamFile.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ExportedData.xlsx");
                    }
                }
            }
        }
    }
}
