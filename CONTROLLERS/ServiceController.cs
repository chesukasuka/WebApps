using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using WebApps.Models;

namespace WebApps.Controllers
{
    public class ServiceController : Controller
    {
        private readonly ILogger<ServiceController> _logger;

        public ServiceController(ILogger<ServiceController> logger)
        {
            _logger = logger;
        }

        public IActionResult Benchmarking()
        {
            return View();
        }

    }
}
