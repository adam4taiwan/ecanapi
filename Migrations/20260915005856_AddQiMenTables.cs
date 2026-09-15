using Microsoft.EntityFrameworkCore.Migrations;
using Npgsql.EntityFrameworkCore.PostgreSQL.Metadata;

#nullable disable

namespace Ecanapi.Migrations
{
    /// <inheritdoc />
    public partial class AddQiMenTables : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<string>(
                name: "DayStem",
                table: "NineStarDailyRules",
                type: "text",
                nullable: false,
                defaultValue: "");

            migrationBuilder.CreateTable(
                name: "QiMenBaMens",
                columns: table => new
                {
                    Id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    Name = table.Column<string>(type: "text", nullable: false),
                    DefaultGong = table.Column<int>(type: "integer", nullable: false),
                    WuXing = table.Column<string>(type: "text", nullable: false),
                    JiXiong = table.Column<string>(type: "text", nullable: false),
                    Desc = table.Column<string>(type: "text", nullable: false),
                    SortOrder = table.Column<int>(type: "integer", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_QiMenBaMens", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "QiMenBaShens",
                columns: table => new
                {
                    Id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    Name = table.Column<string>(type: "text", nullable: false),
                    JiXiong = table.Column<string>(type: "text", nullable: false),
                    Desc = table.Column<string>(type: "text", nullable: false),
                    SortOrder = table.Column<int>(type: "integer", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_QiMenBaShens", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "QiMenDiPanMaps",
                columns: table => new
                {
                    Id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    YinYang = table.Column<string>(type: "text", nullable: false),
                    JuShu = table.Column<int>(type: "integer", nullable: false),
                    GongWei = table.Column<int>(type: "integer", nullable: false),
                    TianGan = table.Column<string>(type: "text", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_QiMenDiPanMaps", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "QiMenJiuXings",
                columns: table => new
                {
                    Id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    Name = table.Column<string>(type: "text", nullable: false),
                    DefaultGong = table.Column<int>(type: "integer", nullable: false),
                    WuXing = table.Column<string>(type: "text", nullable: false),
                    JiXiong = table.Column<string>(type: "text", nullable: false),
                    Desc = table.Column<string>(type: "text", nullable: false),
                    SortOrder = table.Column<int>(type: "integer", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_QiMenJiuXings", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "QiMenJuConfigs",
                columns: table => new
                {
                    Id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    YinYang = table.Column<string>(type: "text", nullable: false),
                    JieQi = table.Column<string>(type: "text", nullable: false),
                    JieQiOrder = table.Column<int>(type: "integer", nullable: false),
                    Yuan = table.Column<string>(type: "text", nullable: false),
                    JuShu = table.Column<int>(type: "integer", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_QiMenJuConfigs", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "QiMenTianGans",
                columns: table => new
                {
                    Id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    TianGan = table.Column<string>(type: "text", nullable: false),
                    Type = table.Column<string>(type: "text", nullable: false),
                    JiXiong = table.Column<string>(type: "text", nullable: false),
                    Desc = table.Column<string>(type: "text", nullable: false),
                    SortOrder = table.Column<int>(type: "integer", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_QiMenTianGans", x => x.Id);
                });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "QiMenBaMens");

            migrationBuilder.DropTable(
                name: "QiMenBaShens");

            migrationBuilder.DropTable(
                name: "QiMenDiPanMaps");

            migrationBuilder.DropTable(
                name: "QiMenJiuXings");

            migrationBuilder.DropTable(
                name: "QiMenJuConfigs");

            migrationBuilder.DropTable(
                name: "QiMenTianGans");

            migrationBuilder.DropColumn(
                name: "DayStem",
                table: "NineStarDailyRules");
        }
    }
}
