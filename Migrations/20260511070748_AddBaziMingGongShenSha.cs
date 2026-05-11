using Microsoft.EntityFrameworkCore.Migrations;
using Npgsql.EntityFrameworkCore.PostgreSQL.Metadata;

#nullable disable

namespace Ecanapi.Migrations
{
    /// <inheritdoc />
    public partial class AddBaziMingGongShenSha : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "BaziMingGongStars",
                columns: table => new
                {
                    Id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    Branch = table.Column<string>(type: "character varying(4)", maxLength: 4, nullable: false),
                    StarName = table.Column<string>(type: "character varying(20)", maxLength: 20, nullable: false),
                    LuckLevel = table.Column<string>(type: "character varying(10)", maxLength: 10, nullable: false),
                    Description = table.Column<string>(type: "character varying(200)", maxLength: 200, nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_BaziMingGongStars", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "BaziShenSha12",
                columns: table => new
                {
                    Id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    Sequence = table.Column<int>(type: "integer", nullable: false),
                    ShenShaName = table.Column<string>(type: "character varying(20)", maxLength: 20, nullable: false),
                    LuckLevel = table.Column<string>(type: "character varying(10)", maxLength: 10, nullable: false),
                    ShortDesc = table.Column<string>(type: "character varying(100)", maxLength: 100, nullable: false),
                    FullDesc = table.Column<string>(type: "character varying(500)", maxLength: 500, nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_BaziShenSha12", x => x.Id);
                });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "BaziMingGongStars");

            migrationBuilder.DropTable(
                name: "BaziShenSha12");
        }
    }
}
