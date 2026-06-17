using Microsoft.EntityFrameworkCore.Migrations;
using Npgsql.EntityFrameworkCore.PostgreSQL.Metadata;

#nullable disable

namespace Ecanapi.Migrations
{
    /// <inheritdoc />
    public partial class AddYiZhuLunMing : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "ig",
                schema: "public",
                columns: table => new
                {
                    id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    code = table.Column<string>(type: "text", nullable: false),
                    name = table.Column<string>(type: "text", nullable: true),
                    description = table.Column<string>(type: "text", nullable: true),
                    desc_one = table.Column<string>(type: "text", nullable: true),
                    desc_two = table.Column<string>(type: "text", nullable: true),
                    desc_three = table.Column<string>(type: "text", nullable: true),
                    desc_four = table.Column<string>(type: "text", nullable: true),
                    desc_five = table.Column<string>(type: "text", nullable: true),
                    desc_six = table.Column<string>(type: "text", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_ig", x => x.id);
                });

            migrationBuilder.CreateTable(
                name: "ig64_six",
                schema: "public",
                columns: table => new
                {
                    id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    ig64 = table.Column<string>(type: "text", nullable: true),
                    one_yao = table.Column<string>(type: "text", nullable: true),
                    two_yao = table.Column<string>(type: "text", nullable: true),
                    three_yao = table.Column<string>(type: "text", nullable: true),
                    four_yao = table.Column<string>(type: "text", nullable: true),
                    five_yao = table.Column<string>(type: "text", nullable: true),
                    six_yao = table.Column<string>(type: "text", nullable: true),
                    RowID = table.Column<int>(type: "integer", nullable: false),
                    gongming = table.Column<string>(type: "text", nullable: true),
                    wuxing = table.Column<string>(type: "text", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_ig64_six", x => x.id);
                });

            migrationBuilder.CreateTable(
                name: "YiZhuLunMings",
                columns: table => new
                {
                    Id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    DayPillar = table.Column<string>(type: "character varying(2)", maxLength: 2, nullable: false),
                    Personality = table.Column<string>(type: "text", nullable: true),
                    Poem = table.Column<string>(type: "text", nullable: true),
                    MonthlyAnalysis = table.Column<string>(type: "text", nullable: true),
                    VoidAnalysis = table.Column<string>(type: "text", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_YiZhuLunMings", x => x.Id);
                });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "ig",
                schema: "public");

            migrationBuilder.DropTable(
                name: "ig64_six",
                schema: "public");

            migrationBuilder.DropTable(
                name: "YiZhuLunMings");
        }
    }
}
