using System;
using Microsoft.EntityFrameworkCore.Migrations;
using Npgsql.EntityFrameworkCore.PostgreSQL.Metadata;

#nullable disable

namespace Ecanapi.Migrations
{
    /// <inheritdoc />
    public partial class AddPointsToUser : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.EnsureSchema(
                name: "public");

            migrationBuilder.AddColumn<string>(
                name: "Email",
                table: "Customers",
                type: "text",
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddColumn<int>(
                name: "points",
                table: "AspNetUsers",
                type: "integer",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.CreateTable(
                name: "十二宮化入十二宮",
                columns: table => new
                {
                    unique_id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    mainstar = table.Column<string>(type: "text", nullable: true),
                    position = table.Column<int>(type: "integer", nullable: true),
                    desc = table.Column<string>(type: "text", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_十二宮化入十二宮", x => x.unique_id);
                });

            migrationBuilder.CreateTable(
                name: "十二宮稱呼",
                columns: table => new
                {
                    unique_id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    name = table.Column<string>(type: "text", nullable: true),
                    desc = table.Column<string>(type: "text", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_十二宮稱呼", x => x.unique_id);
                });

            migrationBuilder.CreateTable(
                name: "十二宮廟旺",
                columns: table => new
                {
                    unique_id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    star = table.Column<string>(type: "text", nullable: true),
                    palace = table.Column<string>(type: "text", nullable: true),
                    light = table.Column<string>(type: "text", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_十二宮廟旺", x => x.unique_id);
                });

            migrationBuilder.CreateTable(
                name: "六十甲子日對時",
                columns: table => new
                {
                    unique_id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    Sky = table.Column<string>(type: "text", nullable: true),
                    Month = table.Column<string>(type: "text", nullable: true),
                    time = table.Column<string>(type: "text", nullable: true),
                    desc = table.Column<string>(type: "text", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_六十甲子日對時", x => x.unique_id);
                });

            migrationBuilder.CreateTable(
                name: "天干星剎",
                columns: table => new
                {
                    unique_id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    KIND = table.Column<string>(type: "text", nullable: true),
                    SKYNO = table.Column<string>(type: "text", nullable: true),
                    TOFLO = table.Column<string>(type: "text", nullable: true),
                    STAR = table.Column<string>(type: "text", nullable: true),
                    DESC = table.Column<string>(type: "text", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_天干星剎", x => x.unique_id);
                });

            migrationBuilder.CreateTable(
                name: "天干陰陽五行",
                columns: table => new
                {
                    unique_id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    sky = table.Column<string>(type: "text", nullable: true),
                    yin_yang = table.Column<string>(type: "text", nullable: true),
                    element = table.Column<string>(type: "text", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_天干陰陽五行", x => x.unique_id);
                });

            migrationBuilder.CreateTable(
                name: "日干對地支",
                columns: table => new
                {
                    unique_id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    KIND = table.Column<string>(type: "text", nullable: true),
                    SKYNO = table.Column<string>(type: "text", nullable: true),
                    TOFLO = table.Column<string>(type: "text", nullable: true),
                    STAR = table.Column<string>(type: "text", nullable: true),
                    DESC = table.Column<string>(type: "text", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_日干對地支", x => x.unique_id);
                });

            migrationBuilder.CreateTable(
                name: "日柱對月支",
                columns: table => new
                {
                    unique_id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    SkyFloor = table.Column<string>(type: "text", nullable: true),
                    position = table.Column<string>(type: "text", nullable: true),
                    desc = table.Column<string>(type: "text", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_日柱對月支", x => x.unique_id);
                });

            migrationBuilder.CreateTable(
                name: "日對時星剎",
                columns: table => new
                {
                    unique_id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    SkyFloor = table.Column<string>(type: "text", nullable: true),
                    position = table.Column<string>(type: "text", nullable: true),
                    desc = table.Column<string>(type: "text", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_日對時星剎", x => x.unique_id);
                });

            migrationBuilder.CreateTable(
                name: "先天四化入十二宮",
                columns: table => new
                {
                    unique_id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    mainstar = table.Column<string>(type: "text", nullable: true),
                    position = table.Column<int>(type: "integer", nullable: true),
                    desc = table.Column<string>(type: "text", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_先天四化入十二宮", x => x.unique_id);
                });

            migrationBuilder.CreateTable(
                name: "地支星剎",
                columns: table => new
                {
                    unique_id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    KIND = table.Column<string>(type: "text", nullable: true),
                    SKYNO = table.Column<string>(type: "text", nullable: true),
                    TOFLO = table.Column<string>(type: "text", nullable: true),
                    STAR = table.Column<string>(type: "text", nullable: true),
                    DESC = table.Column<string>(type: "text", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_地支星剎", x => x.unique_id);
                });

            migrationBuilder.CreateTable(
                name: "地支藏干",
                columns: table => new
                {
                    unique_id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    floor = table.Column<string>(type: "text", nullable: true),
                    sky1 = table.Column<string>(type: "text", nullable: true),
                    sky2 = table.Column<string>(type: "text", nullable: true),
                    sky3 = table.Column<string>(type: "text", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_地支藏干", x => x.unique_id);
                });

            migrationBuilder.CreateTable(
                name: "身主",
                columns: table => new
                {
                    unique_id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    floor = table.Column<string>(type: "text", nullable: true),
                    star = table.Column<string>(type: "text", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_身主", x => x.unique_id);
                });

            migrationBuilder.CreateTable(
                name: "命宮主星",
                columns: table => new
                {
                    unique_id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    star = table.Column<string>(type: "text", nullable: true),
                    star_type = table.Column<string>(type: "text", nullable: true),
                    star_desc = table.Column<string>(type: "text", nullable: true),
                    star_value = table.Column<string>(type: "text", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_命宮主星", x => x.unique_id);
                });

            migrationBuilder.CreateTable(
                name: "易經六十四卦",
                columns: table => new
                {
                    gua_id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    gua_value = table.Column<int>(type: "integer", nullable: true),
                    gua_name = table.Column<string>(type: "text", nullable: true),
                    gua_instruction = table.Column<string>(type: "text", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_易經六十四卦", x => x.gua_id);
                });

            migrationBuilder.CreateTable(
                name: "易經六十四卦分類解說",
                columns: table => new
                {
                    unique_id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    gua_id = table.Column<int>(type: "integer", nullable: true),
                    gua_value = table.Column<int>(type: "integer", nullable: true),
                    gua_desc = table.Column<string>(type: "text", nullable: true),
                    gua_type = table.Column<string>(type: "text", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_易經六十四卦分類解說", x => x.unique_id);
                });

            migrationBuilder.CreateTable(
                name: "星曜狀況",
                columns: table => new
                {
                    unique_id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    star = table.Column<string>(type: "text", nullable: true),
                    desc = table.Column<string>(type: "text", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_星曜狀況", x => x.unique_id);
                });

            migrationBuilder.CreateTable(
                name: "納音",
                columns: table => new
                {
                    unique_id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    gan_zhi = table.Column<string>(type: "text", nullable: true),
                    na_yin = table.Column<string>(type: "text", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_納音", x => x.unique_id);
                });

            migrationBuilder.CreateTable(
                name: "財官總論",
                columns: table => new
                {
                    unique_id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    sky = table.Column<string>(type: "text", nullable: true),
                    desc = table.Column<string>(type: "text", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_財官總論", x => x.unique_id);
                });

            migrationBuilder.CreateTable(
                name: "calendar",
                schema: "public",
                columns: table => new
                {
                    西元年 = table.Column<int>(type: "integer", nullable: false),
                    陽月 = table.Column<int>(type: "integer", nullable: false),
                    陽日 = table.Column<int>(type: "integer", nullable: false),
                    年干支 = table.Column<string>(type: "text", nullable: true),
                    月干支 = table.Column<string>(type: "text", nullable: true),
                    日干支 = table.Column<string>(type: "text", nullable: true),
                    日天干 = table.Column<string>(type: "text", nullable: true),
                    節氣 = table.Column<string>(type: "text", nullable: true),
                    陰曆月 = table.Column<string>(type: "text", nullable: true),
                    陰曆日 = table.Column<string>(type: "text", nullable: true),
                    星期 = table.Column<string>(type: "text", nullable: true),
                    季節 = table.Column<string>(type: "text", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_calendar", x => new { x.西元年, x.陽月, x.陽日 });
                });

            migrationBuilder.CreateTable(
                name: "PointRecords",
                columns: table => new
                {
                    Id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    UserId = table.Column<string>(type: "text", nullable: false),
                    Amount = table.Column<int>(type: "integer", nullable: false),
                    CreatedAt = table.Column<DateTime>(type: "timestamp with time zone", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_PointRecords", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "starstyle",
                columns: table => new
                {
                    unique_id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    mapstar = table.Column<string>(type: "text", nullable: true),
                    mainstar = table.Column<string>(type: "text", nullable: true),
                    position = table.Column<float>(type: "real", nullable: true),
                    gd = table.Column<string>(type: "text", nullable: true),
                    bd = table.Column<string>(type: "text", nullable: true),
                    stardesc = table.Column<string>(type: "text", nullable: true),
                    starbyyear = table.Column<string>(type: "text", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_starstyle", x => x.unique_id);
                });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "十二宮化入十二宮");

            migrationBuilder.DropTable(
                name: "十二宮稱呼");

            migrationBuilder.DropTable(
                name: "十二宮廟旺");

            migrationBuilder.DropTable(
                name: "六十甲子日對時");

            migrationBuilder.DropTable(
                name: "天干星剎");

            migrationBuilder.DropTable(
                name: "天干陰陽五行");

            migrationBuilder.DropTable(
                name: "日干對地支");

            migrationBuilder.DropTable(
                name: "日柱對月支");

            migrationBuilder.DropTable(
                name: "日對時星剎");

            migrationBuilder.DropTable(
                name: "先天四化入十二宮");

            migrationBuilder.DropTable(
                name: "地支星剎");

            migrationBuilder.DropTable(
                name: "地支藏干");

            migrationBuilder.DropTable(
                name: "身主");

            migrationBuilder.DropTable(
                name: "命宮主星");

            migrationBuilder.DropTable(
                name: "易經六十四卦");

            migrationBuilder.DropTable(
                name: "易經六十四卦分類解說");

            migrationBuilder.DropTable(
                name: "星曜狀況");

            migrationBuilder.DropTable(
                name: "納音");

            migrationBuilder.DropTable(
                name: "財官總論");

            migrationBuilder.DropTable(
                name: "calendar",
                schema: "public");

            migrationBuilder.DropTable(
                name: "PointRecords");

            migrationBuilder.DropTable(
                name: "starstyle");

            migrationBuilder.DropColumn(
                name: "Email",
                table: "Customers");

            migrationBuilder.DropColumn(
                name: "points",
                table: "AspNetUsers");
        }
    }
}
