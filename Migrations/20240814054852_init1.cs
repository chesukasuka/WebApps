using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace WebApps.Migrations
{
    /// <inheritdoc />
    public partial class init1 : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "Benchmarking",
                columns: table => new
                {
                    BenchmarkingId = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    KlasifikasiUsaha = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: true),
                    JenisKegiatanUsaha = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: true),
                    NamaPerusahaan = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: true),
                    Negara = table.Column<string>(type: "nvarchar(50)", maxLength: 50, nullable: true),
                    Rasio = table.Column<string>(type: "nvarchar(100)", maxLength: 100, nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Benchmarking", x => x.BenchmarkingId);
                });

            migrationBuilder.CreateTable(
                name: "BenchmarkingTahun",
                columns: table => new
                {
                    BenchmarkingTahunId = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    BenchmarkingId = table.Column<int>(type: "int", nullable: false),
                    Tahun = table.Column<int>(type: "int", nullable: false),
                    Rasio = table.Column<double>(type: "float", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_BenchmarkingTahun", x => x.BenchmarkingTahunId);
                });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "Benchmarking");

            migrationBuilder.DropTable(
                name: "BenchmarkingTahun");
        }
    }
}
