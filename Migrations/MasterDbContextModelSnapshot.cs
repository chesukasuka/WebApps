﻿// <auto-generated />
using System;
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
                    b.Property<int>("BenchmarkingId")
                        .ValueGeneratedOnAdd()
                        .HasColumnType("int");

                    SqlServerPropertyBuilderExtensions.UseIdentityColumn(b.Property<int>("BenchmarkingId"));

                    b.Property<int?>("BenchmarkingModelBenchmarkingId")
                        .HasColumnType("int");

                    b.Property<string>("JenisKegiatanUsaha")
                        .HasMaxLength(100)
                        .HasColumnType("nvarchar(100)");

                    b.Property<string>("KlasifikasiUsaha")
                        .HasMaxLength(100)
                        .HasColumnType("nvarchar(100)");

                    b.Property<string>("NamaPerusahaan")
                        .HasMaxLength(100)
                        .HasColumnType("nvarchar(100)");

                    b.Property<string>("Negara")
                        .HasMaxLength(50)
                        .HasColumnType("nvarchar(50)");

                    b.Property<string>("Rasio")
                        .HasMaxLength(100)
                        .HasColumnType("nvarchar(100)");

                    b.HasKey("BenchmarkingId");

                    b.HasIndex("BenchmarkingModelBenchmarkingId");

                    b.ToTable("Benchmarking");
                });

            modelBuilder.Entity("WebApps.Models.ServiceModel.BenchmarkingTahunModel", b =>
                {
                    b.Property<int>("BenchmarkingTahunId")
                        .ValueGeneratedOnAdd()
                        .HasColumnType("int");

                    SqlServerPropertyBuilderExtensions.UseIdentityColumn(b.Property<int>("BenchmarkingTahunId"));

                    b.Property<int>("BenchmarkingId")
                        .HasColumnType("int");

                    b.Property<double>("Rasio")
                        .HasColumnType("float");

                    b.Property<int>("Tahun")
                        .HasColumnType("int");

                    b.HasKey("BenchmarkingTahunId");

                    b.HasIndex("BenchmarkingId");

                    b.ToTable("BenchmarkingTahun");
                });

            modelBuilder.Entity("WebApps.Models.ServiceModel.BenchmarkingModel", b =>
                {
                    b.HasOne("WebApps.Models.ServiceModel.BenchmarkingModel", null)
                        .WithMany("BenchmarkingModels")
                        .HasForeignKey("BenchmarkingModelBenchmarkingId");
                });

            modelBuilder.Entity("WebApps.Models.ServiceModel.BenchmarkingTahunModel", b =>
                {
                    b.HasOne("WebApps.Models.ServiceModel.BenchmarkingModel", "BenchmarkingModel")
                        .WithMany()
                        .HasForeignKey("BenchmarkingId")
                        .OnDelete(DeleteBehavior.Cascade)
                        .IsRequired();

                    b.Navigation("BenchmarkingModel");
                });

            modelBuilder.Entity("WebApps.Models.ServiceModel.BenchmarkingModel", b =>
                {
                    b.Navigation("BenchmarkingModels");
                });
#pragma warning restore 612, 618
        }
    }
}
