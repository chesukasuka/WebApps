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

        [HttpGet("/Converter")]
        public IActionResult Index()
        {
            Dictionary<string, string>[] data = new Dictionary<string, string>[]
            {
                new Dictionary<string, string>()
                {
                    { "NamaFile", "File1.xlsx" },
                    { "Masa", "12" },
                    { "Tahun", "2024" },
                    { "JumlahRow", "9" }
                },
                new Dictionary<string, string>()
                {
                    { "NamaFile", "File2.xlsx" },
                    { "Masa", "12" },
                    { "Tahun", "2024" },
                    { "JumlahRow", "90" }
                },
                new Dictionary<string, string>()
                {
                    { "NamaFile", "File2.xlsx" },
                    { "Masa", "11" },
                    { "Tahun", "2024" },
                    { "JumlahRow", "50" }
                },
                new Dictionary<string, string>()
                {
                    { "NamaFile", "Excel Terbaru.xlsx" },
                    { "Masa", "10" },
                    { "Tahun", "2024" },
                    { "JumlahRow", "10" }
                },
                new Dictionary<string, string>()
                {
                    { "NamaFile", "Excel 2.xlsx" },
                    { "Masa", "9" },
                    { "Tahun", "2024" },
                    { "JumlahRow", "10" }
                },
                new Dictionary<string, string>()
                {
                    { "NamaFile", "File1.xlsx" },
                    { "Masa", "9" },
                    { "Tahun", "2024" },
                    { "JumlahRow", "10" }
                },
                new Dictionary<string, string>()
                {
                    { "NamaFile", "File1.xlsx" },
                    { "Masa", "8" },
                    { "Tahun", "2024" },
                    { "JumlahRow", "10" }
                },
                new Dictionary<string, string>()
                {
                    { "NamaFile", "File1.xlsx" },
                    { "Masa", "8" },
                    { "Tahun", "2024" },
                    { "JumlahRow", "10" }
                },
                new Dictionary<string, string>()
                {
                    { "NamaFile", "File1.xlsx" },
                    { "Masa", "8" },
                    { "Tahun", "2024" },
                    { "JumlahRow", "10" }
                },
                new Dictionary<string, string>()
                {
                    { "NamaFile", "File1.xlsx" },
                    { "Masa", "7" },
                    { "Tahun", "2024" },
                    { "JumlahRow", "10" }
                },
                new Dictionary<string, string>()
                {
                    { "NamaFile", "File1.xlsx" },
                    { "Masa", "8" },
                    { "Tahun", "2024" },
                    { "JumlahRow", "10" }
                },
                new Dictionary<string, string>()
                {
                    { "NamaFile", "File1.xlsx" },
                    { "Masa", "12" },
                    { "Tahun", "2024" },
                    { "JumlahRow", "10" }
                },
                new Dictionary<string, string>()
                {
                    { "NamaFile", "File1.xlsx" },
                    { "Masa", "8" },
                    { "Tahun", "2024" },
                    { "JumlahRow", "10" }
                }
            };

            ViewBag.Data = data;
            return View("Converter");
        }
    }
}