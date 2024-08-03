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
                .Select(z => new BenchmarkingModel{ jenis_kegiatan_usaha = z.jenis_kegiatan_usaha })
                .Where(z => z.jenis_kegiatan_usaha != null)
                .Distinct()
                .ToList();                

                ViewBag.klasifikasi = _context.Benchmarking
                .Select(z => new BenchmarkingModel{ klasifikasi_usaha = z.klasifikasi_usaha })
                .Where(z => z.klasifikasi_usaha != null)
                .Distinct()
                .ToList();                

                ViewBag.ratio = _context.Benchmarking
                .Select(z => new BenchmarkingModel{ rasio = z.rasio })
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

        public IActionResult Hitung(string rasio, string jenis, string klasifikasi)
        {
            List<BenchmarkingModel> oResult;
            try
            {
                oResult = _context.Benchmarking
                .Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis)
                .ToList();                
            }
            catch (System.Exception)
            {                
                throw;
            }
            return Json(oResult);
        }

        public IActionResult Hitung2(string rasio, string jenis, string klasifikasi)
        {
            List<BenchmarkingResult> oResult = new List<BenchmarkingResult>();
            try
            {
                var oModel = new BenchmarkingResult();
                oModel.keterangan = "Minimal";
                oModel.tahun2019 = _context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Min(z => z.tahun2019);
                oModel.tahun2020 = _context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Min(z => z.tahun2020);
                oModel.tahun2021 = _context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Min(z => z.tahun2021);
                oModel.tahun2022 = _context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Min(z => z.tahun2022);
                oModel.tahun2023 = _context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Min(z => z.tahun2023);
                oModel.tahun2024 = _context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Min(z => z.tahun2024);
                oResult.Add(oModel);

                oModel = new BenchmarkingResult();
                oModel.keterangan = "Kuartil 1";
                oModel.tahun2019 = CalQ1(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2019).ToList());
                oModel.tahun2020 = CalQ1(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2020).ToList());
                oModel.tahun2021 = CalQ1(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2021).ToList());
                oModel.tahun2022 = CalQ1(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2022).ToList());
                oModel.tahun2023 = CalQ1(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2023).ToList());
                oModel.tahun2024 = CalQ1(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2024).ToList());
                oResult.Add(oModel);

                oModel = new BenchmarkingResult();
                oModel.keterangan = "Kuartil 2";
                oModel.tahun2019 = CalQ2(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2019).ToList());
                oModel.tahun2020 = CalQ2(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2020).ToList());
                oModel.tahun2021 = CalQ2(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2021).ToList());
                oModel.tahun2022 = CalQ2(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2022).ToList());
                oModel.tahun2023 = CalQ2(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2023).ToList());
                oModel.tahun2024 = CalQ2(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2024).ToList());
                oResult.Add(oModel);

                oModel = new BenchmarkingResult();
                oModel.keterangan = "Kuartil 3";
                oModel.tahun2019 = CalQ3(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2019).ToList());
                oModel.tahun2020 = CalQ3(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2020).ToList());
                oModel.tahun2021 = CalQ3(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2021).ToList());
                oModel.tahun2022 = CalQ3(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2022).ToList());
                oModel.tahun2023 = CalQ3(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2023).ToList());
                oModel.tahun2024 = CalQ3(_context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Select(z => z.tahun2024).ToList());
                oResult.Add(oModel);

                oModel = new BenchmarkingResult();
                oModel.keterangan = "Maksimum";
                oModel.tahun2019 = _context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Max(z => z.tahun2019);
                oModel.tahun2020 = _context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Max(z => z.tahun2020);
                oModel.tahun2021 = _context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Max(z => z.tahun2021);
                oModel.tahun2022 = _context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Max(z => z.tahun2022);
                oModel.tahun2023 = _context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Max(z => z.tahun2023);
                oModel.tahun2024 = _context.Benchmarking.Where(z => z.rasio == rasio && z.klasifikasi_usaha == klasifikasi && z.jenis_kegiatan_usaha == jenis).Max(z => z.tahun2024);
                oResult.Add(oModel);

            }
            catch (System.Exception)
            {
                throw;
            }
            return Json(oResult);
        }

        public static double CalQ1(List<double> data)
        {
            if (data == null || data.Count == 0)
            {
                throw new InvalidOperationException("Cannot calculate quartile for an empty data set.");
            }

            data.Sort();

            int n = data.Count;
            double q1;

            if (n % 2 == 0)
            {
                q1 = (data[(n / 4) - 1] + data[n / 4]) / 2;
            }
            else
            {
                q1 = data[n / 4];
            }

            return q1;
        }
        public static double CalQ2(List<double> data)
        {
            if (data == null || data.Count == 0)
            {
                throw new InvalidOperationException("Cannot calculate median for an empty data set.");
            }

            data.Sort();

            int n = data.Count;
            double median;

            if (n % 2 == 0)
            {
                // Even number of elements
                median = (data[n / 2 - 1] + data[n / 2]) / 2;
            }
            else
            {
                // Odd number of elements
                median = data[n / 2];
            }

            return median;
        }
        public static double CalQ3(List<double> data)
        {
            if (data == null || data.Count == 0)
            {
                throw new InvalidOperationException("Cannot calculate quartile for an empty data set.");
            }

            data.Sort();

            int n = data.Count;
            double q3;

            if (n % 2 == 0)
            {
                q3 = (data[(3 * n / 4) - 1] + data[3 * n / 4]) / 2;
            }
            else
            {
                q3 = data[(3 * n / 4)];
            }

            return q3;
        }
    }
}
