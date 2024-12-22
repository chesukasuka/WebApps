using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json.Linq;
using Syncfusion.EJ2.FileManager.Base;
using Syncfusion.EJ2.Linq;
using System.Collections;
using System.Diagnostics;
using System.Diagnostics.Metrics;
using System.Globalization;
using System.IO;
using WebApps.Models;
using WebApps.Models.ServiceModel;

namespace WebApps.Controllers
{
    public class ServiceController : Controller
    {
        private readonly ILogger<ServiceController> _logger;
        private readonly MasterDbContext _context;
        private readonly IWebHostEnvironment _env;

        public ServiceController(ILogger<ServiceController> logger, MasterDbContext context, IWebHostEnvironment env)
        {
            _logger = logger;
            _context = context;
            _env = env;
        }

        public IActionResult Benchmarking(string? errMsg = null)
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

                ViewBag.subklasifikasi = _context.Benchmarking
                .Select(z => new BenchmarkingModel { SubKlasifikasiUsaha = z.SubKlasifikasiUsaha, KlasifikasiUsaha = z.KlasifikasiUsaha })
                .Where(z => z.SubKlasifikasiUsaha != null)
                .Distinct()
                .ToList();


                ViewBag.metode = _context.Benchmarking
                .Select(z => new BenchmarkingModel { Metode = z.Metode, SubKlasifikasiUsaha = z.SubKlasifikasiUsaha })
                .Where(z => z.Metode != null)
                .Distinct()
                .ToList();

                ViewBag.ratio = _context.Benchmarking
                .Select(z => new BenchmarkingModel { Rasio = z.Rasio, Metode = z.Metode})
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
            ViewData["errMsg"] = errMsg;
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
                var dataTahun = " and [" + tahun1 + "] != '0.00' ";
                var oLoop = tahun2 - tahun1 + 1;
                var oList = new List<Dictionary<string, object>>();

                var sTahun = "[" + tahun1.ToString() + "]";
                for (int i = tahun1; i < tahun2; i++){
                    sTahun = sTahun + ",[" + (i+1).ToString() + "]";

                    dataTahun = dataTahun + " and [" + (i + 1).ToString() + "]  != '0.00'  ";

                }
                var oListHeader = new List<Dictionary<string, object>>();


                var sql = "";
                sql = ""
                    + " select a.NamaPerusahaan, a.Negara, " + sTahun + "              "
                    + " from benchmarking a                                            "
                    + " left join                                                      "
                    + " (                                                              "
                    + " SELECT                                                         "
                    + " *                                                              "
                    + " FROM                                                           "
                    + "     (SELECT BenchmarkingId, Tahun, Rasio                       "
                    + "     FROM BenchmarkingTahun) as SourceTable                     "
                    + " PIVOT                                                          "
                    + " (                                                              "
                    + "     MAX(Rasio)                                                 "
                    + "     FOR Tahun IN (" + sTahun + ")                              "
                    + " ) as PivotTable                                                "
                    + " ) b                                                            "
                    + " on a.BenchmarkingId=b.BenchmarkingId                           "
                    + " WHERE 1 = 1                                                    "
                    + dataTahun;

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
                            if(result != null)
                            {
                                bool bInsert = true;
                                var oListData = new Dictionary<string, object>();
                                oListData.Add("Nama Perusahaan", result.GetValue(0));
                                oListData.Add("Negara", result.GetValue(1));
                                for (int i = 0; i < oLoop; i++)
                                {
                                    oListData.Add(" " + (tahun1 + i).ToString(), Convert.ToDouble(result.GetValue(i + 2)).ToString("F2"));
                                    if (result.GetValue(i + 2).ToString() == "0")
                                    {
                                        bInsert = false;
                                        continue;
                                    }
                                }
                                if (bInsert)
                                {
                                    oListHeader.Add(oListData);
                                }
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
            
            dynamic oResult;

            //CEK Berapakali hit
            string? hitungCount = HttpContext.Session.GetString("HitungCount");
            string? token = HttpContext.Session.GetString("Token");
            if (string.IsNullOrEmpty(token))
            {
                if (!string.IsNullOrEmpty(hitungCount))
                {
                    int counter = int.Parse(hitungCount);
                    if (counter > 1)
                    {
                        oResult = new
                        {
                            need_login = 1
                        };
                        return Json(oResult);
                    }
                    HttpContext.Session.SetString("HitungCount", (counter+1).ToString());
                }
                else
                {
                    HttpContext.Session.SetString("HitungCount", "2");
                }
            }

            var data = ViewBag.sliderValue;
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

        public List<Dictionary<string, object>> Hitung2Function(string rasio, string jenis, string klasifikasi, int tahun1, int tahun2)
        {
            List<Dictionary<string, object>> oResult = new List<Dictionary<string, object>>();
            var oData = Searchbenchmark(rasio, jenis, klasifikasi, tahun1, tahun2);
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

            return oResult;
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

        static string RemoveFirstWord(string text)
        {
            var words = text.Split(' '); // Pisahkan string menjadi array kata-kata
            if (words.Length > 1)
            {
                return string.Join(" ", words, 1, words.Length - 1); // Gabungkan mulai dari kata kedua
            }
            return string.Empty; // Jika hanya satu kata, kembalikan string kosong
        }

        static string maskName(string name)
        {
            var parts = name.Split(' '); // Pisahkan nama berdasarkan spasi
            return string.Join(" ", parts.Select((part, index) =>
            {
                if (index == 0)
                {
                    // Huruf pertama dibuat kecil, sisanya di-*.
                    return char.ToLower(part[0]) + new string('*', part.Length - 1);
                }
                else
                {
                    // Semua kata lainnya diganti dengan tanda * sepanjang panjang katanya
                    return new string('*', part.Length);
                }
            }));
        }

        public ActionResult GeneratePdf(string rasio, string jenis, string klasifikasi, string subklasifikasi, string metode, int tahun1, int tahun2, double penjualan, double pokokPenjualan, double bebanOperasional, double labaKotor, double labaOperasional, double testedParty, string namaperusahaan)
        {

            var benchmarkingData = Searchbenchmark(rasio, jenis, klasifikasi, tahun1, tahun2);
            var matricData = Hitung2Function(rasio, jenis, klasifikasi, tahun1, tahun2);
            
            // Menyiapkan stream untuk menulis PDF
            MemoryStream workStream = new MemoryStream();
            Document doc = new Document(PageSize.A4, 25, 25, 30, 30);
            PdfWriter writer = PdfWriter.GetInstance(doc, workStream);

            string sHeader = Path.Combine(_env.WebRootPath, "image/layout/1.png");
            string sFooter = Path.Combine(_env.WebRootPath, "image/layout/2.png");
            writer.PageEvent = new PdfWatermarkHelper("Central Data Access", sHeader, sFooter);

            writer.CloseStream = false;

            doc.SetMargins(80f, 80f, 80f, 80f);  // Left, right, top, bottom

            doc.Open();

            ////Logo
            //string imagePath = Path.Combine(_env.WebRootPath, "image/layout/logo.png"); // Sesuaikan dengan path gambar
            //Image logo = Image.GetInstance(imagePath);
            //logo.ScaleAbsolute(150, 30); // Ubah ukuran gambar (width x height)
            //logo.Alignment = Element.ALIGN_LEFT;
            //doc.Add(logo);

            // Judul
            Font titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12);
            Font titleFontItalic = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, Font.ITALIC);

            PdfPTable tableInfoDownload = new PdfPTable(2);
            tableInfoDownload.WidthPercentage = 55;
            tableInfoDownload.SetWidths(new float[] { 11f, 16f });
            tableInfoDownload.DefaultCell.Border = PdfPCell.NO_BORDER;
            tableInfoDownload.DefaultCell.SetLeading(1.5f, 1.5f);
            tableInfoDownload.HorizontalAlignment = Element.ALIGN_RIGHT;

            // Informasi Data Pembanding
            Font textBold = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8);
            Font textBoldWhite = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8, BaseColor.WHITE);
            Font textFont = FontFactory.GetFont(FontFactory.HELVETICA, 8);
            Font textFontWhite = FontFactory.GetFont(FontFactory.HELVETICA, 8, BaseColor.WHITE);
            Font textFontItalic = FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.ITALIC);

            tableInfoDownload.AddCell(new Phrase("Nama Pengguna", textFont));
            if (HttpContext.Session.GetString("Token") != null)
            {
                tableInfoDownload.AddCell(new Phrase($": {HttpContext.Session.GetString("Name")}", textFont));
            }
            else
            {
                tableInfoDownload.AddCell(new Phrase($": Guest", textFont));
            }
            DateTime utcNow = DateTime.UtcNow;

            // Zona waktu Indonesia - WIB
            TimeZoneInfo wibZone = TimeZoneInfo.FindSystemTimeZoneById("SE Asia Standard Time");
            DateTime wibTime = TimeZoneInfo.ConvertTimeFromUtc(utcNow, wibZone);
            tableInfoDownload.AddCell(new Phrase("Waktu Pengunduhan", textFont));
            tableInfoDownload.AddCell(new Phrase($": {wibTime.ToString("f")}", textFont));
            doc.Add(tableInfoDownload);

            Phrase titleText = new Phrase();
            titleText.Add(new Chunk("Benchmarking ", titleFontItalic));
            titleText.Add(new Chunk("Laporan Keuangan", titleFont));
            Paragraph title = new Paragraph(titleText);
            title.SpacingBefore = 20;
            title.Leading = 8 * 1.5f;
            title.Alignment = Element.ALIGN_CENTER;
            doc.Add(title);

            Paragraph subtitle = new Paragraph($"{jenis} {RemoveFirstWord(klasifikasi)}", titleFont);
            subtitle.SpacingBefore = 5;
            subtitle.SpacingAfter = 15;
            subtitle.Leading = 8 * 1.5f;
            subtitle.Alignment = Element.ALIGN_CENTER;
            doc.Add(subtitle);


            PdfPTable tableInfo = new PdfPTable(4);
            tableInfo.WidthPercentage = 100;
            tableInfo.SetWidths(new float[] { 14f, 20f, 14f, 12f });
            tableInfo.DefaultCell.Border = PdfPCell.NO_BORDER;
            tableInfo.DefaultCell.SetLeading(1.5f, 1.5f);

            var headerinfo1 = new PdfPCell(new Phrase("Informasi Data Pembanding", textBold));
            headerinfo1.Colspan = 2;
            headerinfo1.Border = PdfPCell.NO_BORDER;
            tableInfo.AddCell(headerinfo1);
            var headerinfo2 = new PdfPCell(new Phrase("Informasi Perusahaan", textBold));
            headerinfo2.Colspan = 2;
            headerinfo2.Border = PdfPCell.NO_BORDER;
            tableInfo.AddCell(headerinfo2);
            tableInfo.AddCell(new Phrase("Jenis Kegiatan Usaha", textFont));
            tableInfo.AddCell(new Phrase($": {jenis}", textFont));
            tableInfo.AddCell(new Phrase("Nama Perusahaan", textFont));
            tableInfo.AddCell(new Phrase($": {CapitalizeEachWord(namaperusahaan)}", textFont));

            tableInfo.AddCell(new Phrase("Klasifikasi Usaha", textFont));
            tableInfo.AddCell(new Phrase($": {klasifikasi}", textFont));

            tableInfo.AddCell(new Phrase("Penjualan", textFont));
            tableInfo.AddCell(new Phrase($": {penjualan.ToString("N0")}", textFont));

            tableInfo.AddCell(new Phrase("Subklasifikasi Usaha", textFont));
            tableInfo.AddCell(new Phrase($": {subklasifikasi}", textFont));

            tableInfo.AddCell(new Phrase("Harga Pokok Penjualan", textFont));
            tableInfo.AddCell(new Phrase($": {pokokPenjualan.ToString("N0")}", textFont));

            tableInfo.AddCell(new Phrase("Tahun Pajak", textFont));
            tableInfo.AddCell(new Phrase($": {tahun2}", textFont));

            tableInfo.AddCell(new Phrase("Laba Kotor", textFont));
            tableInfo.AddCell(new Phrase($": {labaKotor.ToString("N0")}", textFont));

            tableInfo.AddCell(new Phrase("Rentang", textFont));
            tableInfo.AddCell(new Phrase($": {tahun2-tahun1+1} / Tahun", textFont));

            tableInfo.AddCell(new Phrase("Beban Operasional", textFont));
            tableInfo.AddCell(new Phrase($": {bebanOperasional.ToString("N0")}", textFont));

            tableInfo.AddCell(new Phrase("Metode", textFont));
            tableInfo.AddCell(new Phrase($": {metode}", textFont));

            tableInfo.AddCell(new Phrase("Laba Operasional", textFont));
            tableInfo.AddCell(new Phrase($": {labaOperasional.ToString("N0")}", textFont));

            tableInfo.AddCell(new Phrase("Rasio", textFont));
            tableInfo.AddCell(new Phrase($": {rasio}", textFont));

            tableInfo.AddCell(new Phrase("Tested Party/\nPihak yang diuji (%)", textFont));
            tableInfo.AddCell(new Phrase($": {testedParty.ToString("N2")} %", textFont));

            tableInfo.AddCell(new Phrase("", textFont));
            tableInfo.AddCell(new Phrase("", textFont));

            tableInfo.SpacingAfter = 10;
            doc.Add(tableInfo);

            Paragraph tableTitle = new Paragraph();
            tableTitle.Add(new Chunk("Rentang Kewajaran", textBold));
            tableTitle.Alignment = Element.ALIGN_CENTER;
            tableTitle.SpacingBefore = 5;
            doc.Add(tableTitle);

            string[] tahuns = benchmarkingData.First().Keys.Skip(2).ToArray();
            PdfPTable tableData = new PdfPTable(3 + tahuns.Length);
            tableData.WidthPercentage = 100;
            tableData.SpacingBefore = 5;
            tableData.DefaultCell.VerticalAlignment = Element.ALIGN_MIDDLE;
            List<float> tableDataWidth = new List<float> { 0.5f, 3f, 2f };
            foreach (var tahun in tahuns)
            {
                tableDataWidth.Add(1f);
            }
            tableData.SetWidths(tableDataWidth.ToArray());

            string[] headers = { "No", "Nama Perusahaan Pembanding", "Negara", "Tahun (%)" };
            foreach (var header in headers)
            {
                var headerCell = new PdfPCell(new Phrase(header, textBoldWhite));
                headerCell.HorizontalAlignment = Element.ALIGN_CENTER;
                headerCell.VerticalAlignment = Element.ALIGN_MIDDLE;
                headerCell.BackgroundColor = new BaseColor(System.Drawing.Color.FromArgb(255, 8, 68, 75));

                if (header != "Tahun (%)")
                {
                    headerCell.Rowspan = 2;
                }
                else
                {
                    headerCell.Rowspan = 1;
                    headerCell.Colspan = tahuns.Length;
                }
                tableData.AddCell(headerCell);
            }

            if(benchmarkingData.Count < 1)
            {
                throw new Exception("No Data!");
            }
            foreach(var tahun in tahuns)
            {
                var yearCell = new PdfPCell(new Phrase($"{tahun}", textBoldWhite));
                yearCell.HorizontalAlignment = Element.ALIGN_CENTER;
                yearCell.VerticalAlignment = Element.ALIGN_MIDDLE;
                yearCell.BackgroundColor = new BaseColor(System.Drawing.Color.FromArgb(255, 8, 68, 75));
                //yearCell.Rowspan = 2;
                tableData.AddCell(yearCell);
            }

            int no = 1;
            foreach(var benchmarking in benchmarkingData)
            {                
                var noCell = new PdfPCell(new Phrase(no.ToString(), textFont));
                noCell.HorizontalAlignment = Element.ALIGN_CENTER;
                noCell.VerticalAlignment = Element.ALIGN_MIDDLE;
                tableData.AddCell(noCell);

                var namaPerusahaan = "";
                if (HttpContext.Session.GetString("Token") != null)
                {
                    namaperusahaan = benchmarking.Where(x => x.Key == "Nama Perusahaan").First().Value.ToString();
                }
                else {
                    namaperusahaan = maskName(benchmarking.Where(x => x.Key == "Nama Perusahaan").First().Value.ToString());
                }
                tableData.AddCell(new Phrase(namaperusahaan, textFont));

                var cellNegara = new PdfPCell(new Phrase(benchmarking.Where(x => x.Key == "Negara").First().Value.ToString(), textFont));
                cellNegara.HorizontalAlignment = Element.ALIGN_CENTER;
                cellNegara.VerticalAlignment = Element.ALIGN_MIDDLE;
                tableData.AddCell(cellNegara);

                foreach(var tahun in tahuns)
                {
                    var cell = new PdfPCell(new Phrase(benchmarking.Where(x => x.Key == tahun).First().Value.ToString(), textFont));
                    cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                    cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                    tableData.AddCell(cell);
                }
                no++;
            }

            var tableBreak1 = new PdfPCell(new Paragraph(new Phrase(" ", textFont)));
            tableBreak1.Colspan = 3 + tahuns.Length;
            tableBreak1.HorizontalAlignment = Element.ALIGN_RIGHT;
            tableBreak1.VerticalAlignment = Element.ALIGN_MIDDLE;
            tableBreak1.BackgroundColor = new BaseColor(System.Drawing.Color.FromArgb(255, 8, 68, 75));
            tableData.AddCell(tableBreak1);

            //var tableBreak2 = new PdfPCell(new Phrase("", textFont));
            //tableBreak2.Colspan = tahuns.Length;
            //tableBreak2.BackgroundColor = new BaseColor(System.Drawing.Color.FromArgb(1000, 8, 68, 75));
            //tableData.AddCell(tableBreak2);

            string[] metrics = { "Minimum", "Kuartil 1", "Kuartil 2", "Kuartil 3", "Maksimum" };
            int i = 0;
            foreach (var metric in metrics)
            {
                var metricCell = new PdfPCell(new Phrase(metric, textBold));
                metricCell.Colspan = 3;
                metricCell.HorizontalAlignment = Element.ALIGN_RIGHT;
                metricCell.VerticalAlignment = Element.ALIGN_MIDDLE;
                tableData.AddCell(metricCell);

                foreach(var tahun in tahuns)
                {
                    var metricCellValue = new PdfPCell(new Phrase(matricData[i].Where(x => x.Key == tahun).First().Value.ToString(), textFont));
                    metricCellValue.HorizontalAlignment = Element.ALIGN_RIGHT;
                    metricCellValue.VerticalAlignment = Element.ALIGN_MIDDLE;
                    tableData.AddCell(metricCellValue);
                }
                i++;
            }

            doc.Add(tableData);

            

            // Footer
            Paragraph footer = new Paragraph();
            footer.Add(new Chunk("Data ", textFont));
            footer.Add(new Chunk("benchmarking ", textFontItalic));
            footer.Add(new Chunk("ini ditujukan sebagai informasi umum dan tidak dimaksudkan sebagai rekomendasi spesifik. Untuk penjelasan lebih lanjut mengenai hasil analisis atau perusahaan pembanding, silakan langsung melalui kontak resmi PT Central Data Access (“CDA”).", textFont));
            footer.Alignment = Element.ALIGN_JUSTIFIED;
            footer.SpacingBefore = 5;
            doc.Add(footer);
            
            Paragraph contact = new Paragraph();
            contact.Add(new Chunk("CDA tidak bertanggung jawab atas penggunaan data dan ketidaksesuaian data untuk kebutuhan eksternal kecuali telah melalui tahapan konsultasi resmi melalui ", textFont));
            contact.Add(new Chunk("Hotline ", textFontItalic));
            contact.Add(new Chunk("CDA: 0822-1001-9696.", textFont));
            contact.Alignment = Element.ALIGN_JUSTIFIED;
            contact.SpacingBefore = 5;
            contact.SpacingAfter = 5;
            doc.Add(contact);

            //Logo
            string barcodePath = Path.Combine(_env.WebRootPath, "image/service/barcode.png"); // Sesuaikan dengan path gambar
            Image barcode = Image.GetInstance(barcodePath);
            barcode.ScaleAbsolute(120, 120); // Ubah ukuran gambar (width x height)
            barcode.Alignment = Element.ALIGN_RIGHT;
            doc.Add(barcode);

            // Close Document
            doc.Close();

            byte[] byteInfo = workStream.ToArray();
            workStream.Write(byteInfo, 0, byteInfo.Length);
            workStream.Position = 0;

            Response.Headers.Append("Referrer-Policy", "strict-origin-when-cross-origin");
            // Return PDF as a file result
            return File(workStream, "application/pdf", "CDA_Benchmarking_Laporan_Keuangan.pdf");
        }


        private static string CapitalizeEachWord(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            TextInfo textInfo = CultureInfo.CurrentCulture.TextInfo;

            string Result = textInfo.ToTitleCase(input.ToLower());
            return Result.Replace("Pt ","PT ");
        }
    }
}

public class PdfWatermarkHelper : PdfPageEventHelper
{
    private string _watermarkText;
    private Image _headerImage;
    private Image _footerImage;

    public PdfWatermarkHelper(string watermarkText, string headerImagePath, string footerImagePath)
    {
        _watermarkText = watermarkText;
        _headerImage = Image.GetInstance(headerImagePath);
        _footerImage = Image.GetInstance(footerImagePath);
    }


    public override void OnEndPage(PdfWriter writer, Document document)
    {
        PdfContentByte canvas = writer.DirectContentUnder;
        Font watermarkFont = new Font(Font.FontFamily.HELVETICA, 40, Font.BOLD, new BaseColor(245, 245, 245));
        Phrase watermark = new Phrase(_watermarkText, watermarkFont);

        //// Center the watermark
        //ColumnText.ShowTextAligned(
        //    canvas,
        //    Element.ALIGN_CENTER,
        //    watermark,
        //    document.PageSize.Width / 2,
        //    document.PageSize.Height / 2,
        //    45 // Rotation angle in degrees
        //);


        float fontSize = 50f;
        float opacity = 0.1f;

        // Set the font and opacity
        BaseFont baseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.WINANSI, BaseFont.EMBEDDED);
        canvas.SaveState();
        PdfGState gState = new PdfGState { FillOpacity = opacity };
        canvas.SetGState(gState);
        canvas.SetColorFill(BaseColor.GRAY);

        // Calculate the number of repetitions
        float startX = 0;
        float startY = 0;
        float stepX = 180; // Horizontal spacing
        float stepY = 350; // Vertical spacing

        // Loop to create repeated watermarks
        for (float x = startX; x < PageSize.A4.Width; x += stepX)
        {
            for (float y = startY; y < PageSize.A4.Height; y += stepY)
            {
                // Apply rotation for diagonal text
                canvas.BeginText();
                canvas.SetFontAndSize(baseFont, fontSize);
                canvas.ShowTextAligned(Element.ALIGN_CENTER, _watermarkText, x, y, 45); // Rotate by 45 degrees
                canvas.EndText();
            }
        }

        canvas.RestoreState();

        // Add header image
        float headerWidth = document.PageSize.Width;
        _headerImage.ScaleToFit(headerWidth, _headerImage.Height); // Scale width and preserve aspect ratio
        _headerImage.SetAbsolutePosition(
            0,                                  // Start from the left margin
            document.PageSize.Height - _headerImage.ScaledHeight // Position 10 units from top
        );
        canvas.AddImage(_headerImage);

        // Add footer image
        float footerWidth = document.PageSize.Width;
        _footerImage.ScaleToFit(footerWidth, _footerImage.Height); // Scale width and preserve aspect ratio
        _footerImage.SetAbsolutePosition(
            0, // Start from the left margin
            0                   // Position 10 units from bottom
        );
        canvas.AddImage(_footerImage);
    }

}