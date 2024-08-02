using Microsoft.AspNetCore.Mvc;
using Syncfusion.EJ2.Linq;
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
                // ViewBag.sliderValue = new int[] { UtilityController.dtYear-1, UtilityController.dtYear-1 };

                ViewBag.jenis = _context.Benchmarking
                .Select(z => new BenchmarkingModel { jenis_kegiatan_usaha = z.jenis_kegiatan_usaha })
                .Where(z => z.jenis_kegiatan_usaha != null)
                .Distinct()
                .ToList();

                ViewBag.klasifikasi = _context.Benchmarking
                .Select(z => new BenchmarkingModel { klasifikasi_usaha = z.klasifikasi_usaha })
                .Where(z => z.klasifikasi_usaha != null)
                .Distinct()
                .ToList();

                ViewBag.ratio = _context.Benchmarking
                .Select(z => new BenchmarkingModel { rasio = z.rasio })
                .Where(z => z.rasio != null)
                .Distinct()
                .ToList();

                ViewBag.dataSource = _context.Benchmarking
                .ToList();
                // ViewBag.dataSource = null;

            }
            catch (System.Exception e)
            {
                msg = e.Message;
            }

            return View();
        }

        public class oPayload
        {
            public string? rasio { get; set; }
            public string? jenis { get; set; }
            public string? klasifikasi { get; set; }
        }

        [HttpPost]
        public IActionResult Hitung([FromBody] oPayload payload)
        {
            dynamic oResult;
            try
            {
                oResult = _context.Benchmarking
                .Where(z => z.rasio == payload.rasio && z.klasifikasi_usaha == payload.klasifikasi && z.jenis_kegiatan_usaha == payload.jenis)
                .ToList();

            }
            catch (System.Exception)
            {
                throw;
            }

            return Json(oResult);

        }
    }
}
