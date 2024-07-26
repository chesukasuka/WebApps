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


        public IActionResult test(){
            try
            {
                using (var context = new MasterDbContext())
                {
                    // // Create
                    // var product = new Product { Name = "Product A", Price = 10.99m };
                    // context.Products.Add(product);
                    // context.SaveChanges();

                    // // Read
                    // var products = context.Products.ToList();

                    // // Update
                    // product.Price = 12.99m;
                    // context.SaveChanges();

                    // // Delete
                    // context.Products.Remove(product);
                    // context.SaveChanges();
                }                
            }
            catch (System.Exception)
            {                
                throw;
            }
        }
    }
}
