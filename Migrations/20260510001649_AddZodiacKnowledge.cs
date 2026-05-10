using Microsoft.EntityFrameworkCore.Migrations;
using Npgsql.EntityFrameworkCore.PostgreSQL.Metadata;

#nullable disable

namespace Ecanapi.Migrations
{
    /// <inheritdoc />
    public partial class AddZodiacKnowledge : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "六神四柱口訣",
                columns: table => new
                {
                    id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    position = table.Column<string>(type: "text", nullable: true),
                    star = table.Column<string>(type: "text", nullable: true),
                    simple = table.Column<string>(type: "text", nullable: true),
                    pillar = table.Column<string>(type: "text", nullable: true),
                    gd = table.Column<string>(type: "text", nullable: true),
                    newdesc = table.Column<string>(type: "text", nullable: true),
                    uid = table.Column<int>(type: "integer", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_六神四柱口訣", x => x.id);
                });

            migrationBuilder.CreateTable(
                name: "生肖命理庫",
                columns: table => new
                {
                    id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    birth_branch = table.Column<string>(type: "text", nullable: false),
                    birth_zodiac = table.Column<string>(type: "text", nullable: false),
                    category = table.Column<string>(type: "text", nullable: false),
                    subcategory = table.Column<string>(type: "text", nullable: true),
                    target_branch = table.Column<string>(type: "text", nullable: true),
                    target_label = table.Column<string>(type: "text", nullable: true),
                    content = table.Column<string>(type: "text", nullable: false),
                    fortune_level = table.Column<string>(type: "text", nullable: true),
                    uid = table.Column<int>(type: "integer", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_生肖命理庫", x => x.id);
                });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "六神四柱口訣");

            migrationBuilder.DropTable(
                name: "生肖命理庫");
        }
    }
}
