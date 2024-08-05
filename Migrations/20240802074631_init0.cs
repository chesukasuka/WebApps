using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace WebApps.Migrations
{
    /// <inheritdoc />
    public partial class init0 : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "Benchmarking",
                columns: table => new
                {
                    id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    klasifikasi_usaha = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    jenis_kegiatan_usaha = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    nama_perusahaan = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    negara = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    rasio = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    tahun2019 = table.Column<double>(type: "float", nullable: false),
                    tahun2020 = table.Column<double>(type: "float", nullable: false),
                    tahun2021 = table.Column<double>(type: "float", nullable: false),
                    tahun2022 = table.Column<double>(type: "float", nullable: false),
                    tahun2023 = table.Column<double>(type: "float", nullable: false),
                    tahun2024 = table.Column<double>(type: "float", nullable: false),
                    rasio2025 = table.Column<double>(type: "float", nullable: false),
                    rasio2026 = table.Column<double>(type: "float", nullable: false),
                    rasio2027 = table.Column<double>(type: "float", nullable: false),
                    rasio2028 = table.Column<double>(type: "float", nullable: false),
                    rasio2029 = table.Column<double>(type: "float", nullable: false),
                    rasio2030 = table.Column<double>(type: "float", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Benchmarking", x => x.id);
                });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "Benchmarking");
        }
    }
}
