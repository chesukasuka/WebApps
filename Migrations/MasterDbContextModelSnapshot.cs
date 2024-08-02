﻿// <auto-generated />
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Infrastructure;
using Microsoft.EntityFrameworkCore.Metadata;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;
using WebApps.Controllers;

#nullable disable

namespace WebApps.Migrations
{
    [DbContext(typeof(MasterDbContext))]
    partial class MasterDbContextModelSnapshot : ModelSnapshot
    {
        protected override void BuildModel(ModelBuilder modelBuilder)
        {
#pragma warning disable 612, 618
            modelBuilder
                .HasAnnotation("ProductVersion", "8.0.7")
                .HasAnnotation("Relational:MaxIdentifierLength", 128);

            SqlServerModelBuilderExtensions.UseIdentityColumns(modelBuilder);

            modelBuilder.Entity("WebApps.Models.ServiceModel.BenchmarkingModel", b =>
                {
                    b.Property<int>("id")
                        .ValueGeneratedOnAdd()
                        .HasColumnType("int");

                    SqlServerPropertyBuilderExtensions.UseIdentityColumn(b.Property<int>("id"));

                    b.Property<string>("jenis_kegiatan_usaha")
                        .HasColumnType("nvarchar(max)");

                    b.Property<string>("klasifikasi_usaha")
                        .HasColumnType("nvarchar(max)");

                    b.Property<string>("nama_perusahaan")
                        .HasColumnType("nvarchar(max)");

                    b.Property<string>("negara")
                        .HasColumnType("nvarchar(max)");

                    b.Property<string>("rasio")
                        .HasColumnType("nvarchar(max)");

                    b.Property<double>("rasio2025")
                        .HasColumnType("float");

                    b.Property<double>("rasio2026")
                        .HasColumnType("float");

                    b.Property<double>("rasio2027")
                        .HasColumnType("float");

                    b.Property<double>("rasio2028")
                        .HasColumnType("float");

                    b.Property<double>("rasio2029")
                        .HasColumnType("float");

                    b.Property<double>("rasio2030")
                        .HasColumnType("float");

                    b.Property<double>("tahun2019")
                        .HasColumnType("float");

                    b.Property<double>("tahun2020")
                        .HasColumnType("float");

                    b.Property<double>("tahun2021")
                        .HasColumnType("float");

                    b.Property<double>("tahun2022")
                        .HasColumnType("float");

                    b.Property<double>("tahun2023")
                        .HasColumnType("float");

                    b.Property<double>("tahun2024")
                        .HasColumnType("float");

                    b.HasKey("id");

                    b.ToTable("Benchmarking");
                });
#pragma warning restore 612, 618
        }
    }
}
