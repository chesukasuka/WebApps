using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json.Linq;
using Syncfusion.EJ2.Linq;
using System.Collections;
using System.Diagnostics;
using System.Diagnostics.Metrics;
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

                ViewBag.metode = _context.Benchmarking
                .Select(z => new BenchmarkingModel { Metode = z.Metode, KlasifikasiUsaha = z.KlasifikasiUsaha })
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

        public ActionResult GeneratePdf(string rasio, string jenis, string klasifikasi, int tahun1, int tahun2, double penjualan, double pokokPenjualan, double bebanOperasional, double labaKotor, double labaOperasional, double testedParty, string namaperusahaan)
        {
            var benchmarkingData = Searchbenchmark(rasio, jenis, klasifikasi, tahun1, tahun2);
            var matricData = Hitung2Function(rasio, jenis, klasifikasi, tahun1, tahun2);
            
            // Menyiapkan stream untuk menulis PDF
            MemoryStream workStream = new MemoryStream();
            Document doc = new Document(PageSize.A4, 25, 25, 30, 30);
            PdfWriter writer = PdfWriter.GetInstance(doc, workStream);

            writer.PageEvent = new PdfWatermarkHelper("Central Data Access");

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

            Paragraph subtitle = new Paragraph($"{jenis} {RemoveFirstWord(klasifikasi)}\n{DateTime.Now.ToString("dd MMMM yyyy")}", titleFont);
            subtitle.SpacingAfter = 20;
            subtitle.Leading = 8 * 1.5f;
            subtitle.Alignment = Element.ALIGN_CENTER;
            doc.Add(subtitle);

            // Informasi Data Pembanding
            Font textBold = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 8);
            Font textFont = FontFactory.GetFont(FontFactory.HELVETICA, 8);

            PdfPTable tableInfo = new PdfPTable(4);
            tableInfo.WidthPercentage = 100;
            tableInfo.SetWidths(new float[] { 14f, 16f, 14f, 16f });
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
            tableInfo.AddCell(new Phrase($": {namaperusahaan}", textFont));

            tableInfo.AddCell(new Phrase("Klasifikasi Usaha", textFont));
            tableInfo.AddCell(new Phrase($": {klasifikasi}", textFont));
            tableInfo.AddCell(new Phrase("Penjualan", textFont));
            tableInfo.AddCell(new Phrase($": {penjualan.ToString("N0")}", textFont));
            //tableInfo.AddCell(new Phrase("Subklasifikasi Usaha", textFont));
            //tableInfo.AddCell(new Phrase(":", textFont));
            tableInfo.AddCell(new Phrase("Tahun Pajak", textFont));
            tableInfo.AddCell(new Phrase($": {tahun2}", textFont));
            tableInfo.AddCell(new Phrase("Harga Pokok Pendapatan", textFont));
            tableInfo.AddCell(new Phrase($": {labaKotor.ToString("N0")}", textFont));

            tableInfo.AddCell(new Phrase("Rasio Keuangan", textFont));
            tableInfo.AddCell(new Phrase($": {rasio}", textFont));
            tableInfo.AddCell(new Phrase("Beban Operasional", textFont));
            tableInfo.AddCell(new Phrase($": {bebanOperasional.ToString("N0")}", textFont));

            tableInfo.AddCell(new Phrase("", textFont));
            tableInfo.AddCell(new Phrase("", textFont));
            tableInfo.AddCell(new Phrase("Laba Operasional", textFont));
            tableInfo.AddCell(new Phrase($": {labaOperasional.ToString("N0")}", textFont));

            doc.Add(tableInfo);

            
            string[] tahuns = benchmarkingData.First().Keys.Skip(2).ToArray();
            PdfPTable tableData = new PdfPTable(3 + tahuns.Length);
            tableData.WidthPercentage = 100;
            tableData.SpacingBefore = 30;
            tableData.DefaultCell.VerticalAlignment = Element.ALIGN_MIDDLE;
            List<float> tableDataWidth = new List<float> { 0.5f, 3f, 2f };
            foreach (var tahun in tahuns)
            {
                tableDataWidth.Add(1f);
            }
            tableData.SetWidths(tableDataWidth.ToArray());

            string[] headers = { "No", "Perusahaan", "Negara" };
            foreach (var header in headers)
            {
                var headerCell = new PdfPCell(new Phrase(header, textBold));
                headerCell.HorizontalAlignment = Element.ALIGN_CENTER;
                headerCell.BackgroundColor = new BaseColor(System.Drawing.Color.LightBlue);

                //if (header != "NCPM (%)")
                //{
                //    headerCell.Rowspan = 2;
                //}
                //else
                //{
                //    headerCell.Rowspan = 1;
                //    headerCell.Colspan = 3;
                //}
                headerCell.Rowspan = 2;
                tableData.AddCell(headerCell);
            }

            if(benchmarkingData.Count < 1)
            {
                throw new Exception("No Data!");
            }
            foreach(var tahun in tahuns)
            {
                var yearCell = new PdfPCell(new Phrase(tahun, textBold));
                yearCell.HorizontalAlignment = Element.ALIGN_CENTER;
                yearCell.BackgroundColor = new BaseColor(System.Drawing.Color.LightBlue);
                yearCell.Rowspan = 2;
                tableData.AddCell(yearCell);
            }

            int no = 1;
            foreach(var benchmarking in benchmarkingData)
            {                
                var noCell = new PdfPCell(new Phrase(no.ToString(), textFont));
                noCell.HorizontalAlignment = Element.ALIGN_CENTER;
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
                tableData.AddCell(new Phrase(benchmarking.Where(x => x.Key == "Negara").First().Value.ToString(), textFont));
                foreach(var tahun in tahuns)
                {
                    var cell = new PdfPCell(new Phrase(benchmarking.Where(x => x.Key == tahun).First().Value.ToString(), textFont));
                    cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                    tableData.AddCell(cell);
                }
                no++;
            }

            var tableBreak1 = new PdfPCell(new Phrase("Rasio Keuangan Perusahaan", textFont));
            tableBreak1.Colspan = 3;
            tableBreak1.HorizontalAlignment = Element.ALIGN_RIGHT;
            tableBreak1.BackgroundColor = new BaseColor(System.Drawing.Color.LightBlue);
            tableData.AddCell(tableBreak1);

            var tableBreak2 = new PdfPCell(new Phrase("", textFont));
            tableBreak2.Colspan = tahuns.Length;
            tableBreak2.BackgroundColor = new BaseColor(System.Drawing.Color.LightBlue);
            tableData.AddCell(tableBreak2);

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
            Paragraph footer = new Paragraph(
                "Dokumen Penetapan Harga Transfer menguji kewajaran transaksi afiliasi yang dilakukan dengan mencari perusahaan pembanding yang sering disebut dengan Benchmarking Data Pembanding. Data Pembanding Eksternal yang tersedia diperoleh dari BvD TPCatalyst yang kemudian diolah menjadi Benchmarking Data Pembanding yang di klasifikasikan berdasarkan klasifikasi usaha." +
                "\n\n" +
                "Berdasarkan data pembanding milik Central Data Access, batas kewajaran harga transfer sesuai dengan Prinsip Kewajaran dan Kelaziman Usaha, margin penentuan harga transfer yang ditetapkan sebaiknya berada dalam pertimbangan rentang kewajaran.",                
                textFont
            );
            footer.Alignment = Element.ALIGN_JUSTIFIED;
            footer.SpacingBefore = 20;
            doc.Add(footer);
            
            Paragraph contact = new Paragraph(
                "Konsultasi Gratis bersama Tim Central Data Access untuk mendapatkan analisis data komprehensif terkait dengan Dokumen Penentuan Harga Transfer" +
                "\n\n" +
                "www.centraldataaccess.co.id | Telepon : (0251) 8574 375 | Handphone : 0822 1001 9696 / 0822 1001 9797 | \r\ncs@centraldataaccess.co.id",
                textBold
            );
            contact.Alignment = Element.ALIGN_CENTER;
            contact.SpacingBefore = 40;
            doc.Add(contact);

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

public class PdfWatermarkHelper : PdfPageEventHelper
{
    private string _watermarkText;

    public PdfWatermarkHelper(string watermarkText)
    {
        _watermarkText = watermarkText;
    }

    public override void OnEndPage(PdfWriter writer, Document document)
    {
        PdfContentByte canvas = writer.DirectContentUnder;
        Font watermarkFont = new Font(Font.FontFamily.HELVETICA, 60, Font.BOLD, new BaseColor(200, 200, 200));
        Phrase watermark = new Phrase(_watermarkText, watermarkFont);

        // Center the watermark
        ColumnText.ShowTextAligned(
            canvas,
            Element.ALIGN_CENTER,
            watermark,
            document.PageSize.Width / 2,
            document.PageSize.Height / 2,
            45 // Rotation angle in degrees
        );
    }
}