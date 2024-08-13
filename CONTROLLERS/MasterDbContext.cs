using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Diagnostics;
using WebApps.Models;
using WebApps.Models.ServiceModel;

namespace WebApps.Controllers
{
    public class MasterDbContext : DbContext
    {
        public MasterDbContext(DbContextOptions<MasterDbContext> options) : base(options) { }

        public DbSet<BenchmarkingModel> Benchmarking { get; set; }
        public DbSet<BenchmarkingTahunModel> BenchmarkingTahun { get; set; }
    }
}