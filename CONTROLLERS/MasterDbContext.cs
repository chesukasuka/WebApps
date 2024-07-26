using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using WebApps.Models;

namespace WebApps.Controllers
{
    public class MasterDbContext : DbContext
    {
        public MasterDbContext(DbContextOptions<MasterDbContext> options) : base(options) { }

        public DbSet<BenchmarkingModel> Benchmarking { get; set; }
    }
}