using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.AspNetCore.Hosting.Server;
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
        private readonly IWebHostEnvironment _env;

        public ServiceController(ILogger<ServiceController> logger, MasterDbContext context, IWebHostEnvironment env)
        {
            _logger = logger;
            _context = context;
            _env = env;
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

        public ActionResult GeneratePdf()
        {
            // Menyiapkan stream untuk menulis PDF
            MemoryStream workStream = new MemoryStream();
            Document doc = new Document(PageSize.A4, 25, 25, 30, 30);
            PdfWriter writer = PdfWriter.GetInstance(doc, workStream);
            writer.CloseStream = false;

            doc.SetMargins(80f, 80f, 80f, 80f);  // Left, right, top, bottom

            doc.Open();

            //Logo
            string imagePath = Path.Combine(_env.WebRootPath, "image/layout/logo.png"); // Sesuaikan dengan path gambar
            Image logo = Image.GetInstance(imagePath);
            logo.ScaleAbsolute(150, 30); // Ubah ukuran gambar (width x height)
            logo.Alignment = Element.ALIGN_LEFT;
            doc.Add(logo);

            // Judul
            Font titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8);
            Font titleFontItalic = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8, Font.ITALIC);
            Phrase titleText = new Phrase();
            titleText.Add(new Chunk("Benchmarking ", titleFontItalic));
            titleText.Add(new Chunk("Laporan Keuangan", titleFont));
            Paragraph title = new Paragraph(titleText);
            title.SpacingBefore = 25;
            title.Leading = 8 * 1.5f;
            title.Alignment = Element.ALIGN_CENTER;
            doc.Add(title);

            Paragraph subtitle = new Paragraph("Distributor Alat Kesehatan", titleFont);
            subtitle.SpacingAfter = 20;
            subtitle.Leading = 8 * 1.5f;
            subtitle.Alignment = Element.ALIGN_CENTER;
            doc.Add(subtitle);

            // Informasi Data Pembanding
            Font headerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8);
            Font textFont = FontFactory.GetFont(FontFactory.HELVETICA, 10);

            PdfPTable tableInfo = new PdfPTable(4);
            tableInfo.WidthPercentage = 100;
            tableInfo.SetWidths(new float[] { 20f, 10f, 20f, 10f });
            tableInfo.DefaultCell.Border = PdfPCell.NO_BORDER;
            tableInfo.DefaultCell.SetLeading(1.5f, 1.5f);

            var headerinfo1 = new PdfPCell(new Phrase("Informasi Data Pembanding", headerFont));
            headerinfo1.Colspan = 2;
            headerinfo1.Border = PdfPCell.NO_BORDER;
            tableInfo.AddCell(headerinfo1);
            var headerinfo2 = new PdfPCell(new Phrase("Ringkasan Laporan Keuangan", headerFont));
            headerinfo2.Colspan = 2;
            headerinfo2.Border = PdfPCell.NO_BORDER;
            tableInfo.AddCell(headerinfo2);
            tableInfo.AddCell(new Phrase("Jenis Kegiatan Usaha", textFont));
            tableInfo.AddCell(new Phrase(":", textFont));
            tableInfo.AddCell(new Phrase("Nama Perusahaan", textFont));
            tableInfo.AddCell(new Phrase(":", textFont));
            tableInfo.AddCell(new Phrase("Klasifikasi Usaha", textFont));
            tableInfo.AddCell(new Phrase(":", textFont));
            tableInfo.AddCell(new Phrase("Penjualan", textFont));
            tableInfo.AddCell(new Phrase(":", textFont));
            tableInfo.AddCell(new Phrase("Subklasifikasi Usaha", textFont));
            tableInfo.AddCell(new Phrase(":", textFont));
            tableInfo.AddCell(new Phrase("Harga Pokok Pendapatan", textFont));
            tableInfo.AddCell(new Phrase(":", textFont));
            tableInfo.AddCell(new Phrase("Tahun Pajak", textFont));
            tableInfo.AddCell(new Phrase(":", textFont));
            tableInfo.AddCell(new Phrase("Beban Operasional", textFont));
            tableInfo.AddCell(new Phrase(":", textFont));
            tableInfo.AddCell(new Phrase("Rasio Keuangan", textFont));
            tableInfo.AddCell(new Phrase(":", textFont));
            tableInfo.AddCell(new Phrase("Laba Operasional", textFont));
            tableInfo.AddCell(new Phrase(":", textFont));

            doc.Add(tableInfo);

            // Rasio Keuangan Perusahaan (Table)
            Paragraph sectionTitle = new Paragraph("Rasio Keuangan Perusahaan", headerFont);
            sectionTitle.SpacingBefore = 15;
            sectionTitle.SpacingAfter = 5;
            doc.Add(sectionTitle);

            PdfPTable tableData = new PdfPTable(6);
            tableData.WidthPercentage = 100;
            tableData.SetWidths(new float[] { 0.5f, 2, 2, 2, 2, 2 });

            string[] headers = { "No", "Perusahaan", "Negara", "NCPM (%)" };
            foreach (var header in headers)
            {
                var headerCell = new PdfPCell(new Phrase(header, headerFont));
                if (header != "NCPM (%)")
                {
                    headerCell.Rowspan = 2;
                }
                else
                {
                    headerCell.Rowspan = 1;
                    headerCell.Colspan = 3;
                }
                tableData.AddCell(headerCell);
            }
            string[] tahuns = { "2019", "2020", "2021" };
            foreach(var tahun in tahuns)
            {
                tableData.AddCell(new PdfPCell(new Phrase(tahun, headerFont)));
            }

            string[,] rows = {
                { "1", "K.M.R Co.,Ltd", "Republic of Korea", "", "", "" },
                { "2", "Grepcor, Inc.", "Philippines", "", "", "" },
                { "3", "Veiva Scientific India Private Limited", "India", "", "", "" },
                { "4", "Shin Ki Commercial Co.,Ltd", "Republic of Korea", "", "", "" },
                { "5", "Wooree Technologies Co.,Ltd", "Republic of Korea", "", "", "" },
                { "6", "Insol Co.,Ltd", "Republic of Korea", "", "", "" }
            };

            for (int i = 0; i < rows.GetLength(0); i++)
            {
                for (int j = 0; j < rows.GetLength(1); j++)
                {
                    tableData.AddCell(new Phrase(rows[i, j], textFont));
                }
            }

            doc.Add(tableData);

            // Summary Section
            Paragraph summaryTitle = new Paragraph("Rasio Keuangan Perusahaan", headerFont);
            summaryTitle.SpacingBefore = 15;
            doc.Add(summaryTitle);

            string[] metrics = { "Minimum", "Kuartil 1", "Kuartil 2", "Kuartil 3", "Maksimum" };
            PdfPTable tableMetrics = new PdfPTable(1);
            tableMetrics.WidthPercentage = 100;
            foreach (var metric in metrics)
            {
                tableMetrics.AddCell(new PdfPCell(new Phrase(metric, textFont)));
            }

            doc.Add(tableMetrics);

            // Footer
            Paragraph footer = new Paragraph(
                "Konsultasi Gratis bersama Tim Central Data Access untuk mendapatkan analisis data komprehensif terkait dengan Dokumen Penentuan Harga Transfer\n" +
                "www.centraldataaccess.co.id | Telepon : (0251) 8574 375 | Handphone : 0822 1001 9696 / 0822 1001 9797 | cs@centraldataaccess.co.id",
                textFont
            );
            footer.Alignment = Element.ALIGN_CENTER;
            footer.SpacingBefore = 20;
            doc.Add(footer);

            // Close Document
            doc.Close();

            byte[] byteInfo = workStream.ToArray();
            workStream.Write(byteInfo, 0, byteInfo.Length);
            workStream.Position = 0;

            // Return PDF as a file result
            return File(workStream, "application/pdf", "Benchmarking_Laporan_Keuangan.pdf");
        }

    }
}
