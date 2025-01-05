using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json.Linq;
using Syncfusion.EJ2.FileManager.Base;
using Syncfusion.EJ2.Linq;
using Syncfusion.XlsIO;
using System.IO;
using System.Xml.Linq;
using WebApps.Models;
using WebApps.Models.ServiceModel;

namespace WebApps.Controllers
{
    public class ConverterController : Controller
    {
        private readonly ILogger<ServiceController> _logger;
        private readonly MasterDbContext _context;
        private readonly IWebHostEnvironment _env;

        public ConverterController(ILogger<ServiceController> logger, MasterDbContext context, IWebHostEnvironment env)
        {
            _logger = logger;
            _context = context;
            _env = env;
        }

        public class MappPajak
        {
            public string Baris { get; set; }
            public long? Id { get; set; }
        }

        [HttpGet("/Converter")]
        public IActionResult Index()
        {
            ViewBag.Daftar = _context.FakturKeluaranDaftar
                .Select(z => new FakturKeluaranDaftarModel {
                    FakturKeluaranDaftarId = z.FakturKeluaranDaftarId,
                    Jumlah = z.Jumlah,
                    NamaFile = z.NamaFile,
                })
                .OrderByDescending(z => z.FakturKeluaranDaftarId)
                .ToList();

            return View("Converter");
        }

        [HttpPost("Import")]
        public IActionResult Import(IFormFile file)
        {

            var filePath = @"C:\a\1.xlsx";
            byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);

            using (var transaction = _context.Database.BeginTransaction())
            {
                try
                {
                    using (var stream = new MemoryStream(fileBytes))
                    {
                        using (ExcelEngine excelEngine = new ExcelEngine())
                        {
                            IApplication application = excelEngine.Excel;
                            application.DefaultVersion = ExcelVersion.Excel2016;

                            stream.Position = 0;
                            IWorkbook workbook = application.Workbooks.Open(stream);

                            var ListHeader = new List<MappPajak>();

                            IWorksheet worksheetHeader = workbook.Worksheets["Faktur"];
                            var oModelDaftar = new FakturKeluaranDaftarModel();
                            oModelDaftar.Jumlah = worksheetHeader.UsedRange.LastRow - 4;
                            oModelDaftar.NamaFile = filePath;
                            _context.FakturKeluaranDaftar.Add(oModelDaftar);
                            _context.SaveChanges();
                            var oDaftarId = oModelDaftar.FakturKeluaranDaftarId;

                            for (int row = 1; row <= worksheetHeader.UsedRange.LastRow; row++)
                            {
                                if (row >= 4)
                                {
                                    var oMapPajak = new MappPajak();
                                    var oModelData = new FakturKeluaranHeaderModel();
                                    var oCellValue = new List<String>();
                                    for (int col = 1; col <= worksheetHeader.UsedRange.LastColumn; col++)
                                    {
                                        var cellValue = worksheetHeader[row, col].DisplayText;
                                        oCellValue.Add(cellValue);
                                    }

                                    if (oCellValue[0] == "END")
                                    {
                                        break;
                                    }

                                    oModelData.TaxInvoiceDate = oCellValue[1];
                                    oModelData.TaxInvoiceOpt = oCellValue[2];
                                    oModelData.TrxCode = oCellValue[3];
                                    oModelData.AddInfo = oCellValue[4];
                                    oModelData.CustomDoc = oCellValue[5];
                                    oModelData.RefDesc = oCellValue[6];
                                    oModelData.FacilityStamp = oCellValue[7];
                                    oModelData.SellerIDTKU = oCellValue[8];
                                    oModelData.BuyerTin = oCellValue[9];
                                    oModelData.BuyerDocument = oCellValue[10];
                                    oModelData.BuyerCountry = oCellValue[11];
                                    oModelData.BuyerDocumentNumber = oCellValue[12];
                                    oModelData.BuyerName = oCellValue[13];
                                    oModelData.BuyerAdress = oCellValue[14];
                                    oModelData.BuyerEmail = oCellValue[15];
                                    oModelData.BuyerIDTKU = oCellValue[16];
                                    oModelData.FakturKeluaranDaftarId = oDaftarId;

                                    _context.FakturKeluaranHeader.Add(oModelData);
                                    _context.SaveChanges();

                                    oMapPajak.Baris = oCellValue[0];
                                    oMapPajak.Id = oModelData.FakturKeluaranHeaderId;

                                    ListHeader.Add(oMapPajak);

                                }
                            }

                            IWorksheet worksheetItem = workbook.Worksheets["DetailFaktur"];
                            for (int row = 1; row <= worksheetItem.UsedRange.LastRow; row++)
                            {
                                if (row >= 2)
                                {
                                    var oModelData = new FakturKeluaranItemModel();
                                    var oCellValue = new List<String>();
                                    for (int col = 1; col <= worksheetItem.UsedRange.LastColumn; col++)
                                    {
                                        var cellValue = worksheetItem[row, col].DisplayText;
                                        oCellValue.Add(cellValue);
                                    }

                                    long oId = -1;
                                    foreach (MappPajak oMapPajak in ListHeader)
                                    {
                                        if (oCellValue[0] == oMapPajak.Baris)
                                        {
                                            oId = (long)oMapPajak.Id;
                                            break;
                                        }
                                    }

                                    if (oCellValue[0] == "END")
                                    {
                                        break;
                                    }

                                    oModelData.FakturKeluaranHeaderId = oId;
                                    oModelData.Opt = oCellValue[1];
                                    oModelData.Code = oCellValue[2];
                                    oModelData.Name = oCellValue[3];
                                    oModelData.Unit = oCellValue[4];
                                    oModelData.Price = Convert.ToDouble(oCellValue[5]);
                                    oModelData.Qty = Convert.ToDouble(oCellValue[6]);
                                    oModelData.TotalDiscount = Convert.ToDouble(oCellValue[7]);
                                    oModelData.TaxBase = Convert.ToDouble(oCellValue[8]);
                                    oModelData.OtherTaxBase = Convert.ToDouble(oCellValue[9]);
                                    oModelData.VATRate = oCellValue[10];
                                    oModelData.VAT = Convert.ToDouble(oCellValue[11]);
                                    oModelData.STLGRate = oCellValue[12];
                                    oModelData.STLG = Convert.ToDouble(oCellValue[13]);

                                    _context.FakturKeluaranItem.Add(oModelData);
                                    _context.SaveChanges();
                                }
                            }

                            workbook.Close();
                        }

                    }

                    transaction.Commit();

                }
                catch (Exception ex)
                {
                    transaction.Rollback();
                    Console.WriteLine($"Transaction failed: {ex.Message}");
                }
            }

            return View("Converter");
        }

        [HttpPost("Export")]
        public IActionResult Export(string PajakKeluaranDaftarId)
        {
            try
            {
                XDocument xmlDocument = new XDocument(
                    new XElement("Products", // Root element
                        new XElement("Product",
                            new XAttribute("Id", 1),
                            new XElement("Name", "Product 1"),
                            new XElement("Price", 9.99)
                        ),
                        new XElement("Product",
                            new XAttribute("Id", 2),
                            new XElement("Name", "Product 2"),
                            new XElement("Price", 14.99)
                        )
                    )
                );

                string filePath = "Products.xml";
                xmlDocument.Save(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Transaction failed: {ex.Message}");
            }

            return View("Converter");
        }

    }
}