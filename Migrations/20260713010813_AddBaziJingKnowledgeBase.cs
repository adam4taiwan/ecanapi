using Microsoft.EntityFrameworkCore.Migrations;
using Npgsql.EntityFrameworkCore.PostgreSQL.Metadata;

#nullable disable

namespace Ecanapi.Migrations
{
    /// <inheritdoc />
    public partial class AddBaziJingKnowledgeBase : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<int>(
                name: "Id",
                schema: "public",
                table: "calendar",
                type: "integer",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.CreateTable(
                name: "BaziJingCaiGuan",
                columns: table => new
                {
                    Id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    Category = table.Column<string>(type: "text", nullable: false),
                    ConfigType = table.Column<string>(type: "text", nullable: false),
                    Condition = table.Column<string>(type: "text", nullable: true),
                    Content = table.Column<string>(type: "text", nullable: true),
                    SortOrder = table.Column<int>(type: "integer", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_BaziJingCaiGuan", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "BaziJingConfig",
                columns: table => new
                {
                    Id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    ConfigType = table.Column<string>(type: "text", nullable: false),
                    ConfigName = table.Column<string>(type: "text", nullable: false),
                    Condition = table.Column<string>(type: "text", nullable: true),
                    Content = table.Column<string>(type: "text", nullable: true),
                    SortOrder = table.Column<int>(type: "integer", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_BaziJingConfig", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "BaziJingKouJue",
                columns: table => new
                {
                    Id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    Category = table.Column<string>(type: "text", nullable: false),
                    Condition = table.Column<string>(type: "text", nullable: true),
                    Content = table.Column<string>(type: "text", nullable: true),
                    SortOrder = table.Column<int>(type: "integer", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_BaziJingKouJue", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "BaziJingLiuQin",
                columns: table => new
                {
                    Id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    LiuQinType = table.Column<string>(type: "text", nullable: false),
                    Category = table.Column<string>(type: "text", nullable: false),
                    Condition = table.Column<string>(type: "text", nullable: true),
                    Content = table.Column<string>(type: "text", nullable: true),
                    SortOrder = table.Column<int>(type: "integer", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_BaziJingLiuQin", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "BaziJingShenSha",
                columns: table => new
                {
                    Id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    Name = table.Column<string>(type: "text", nullable: false),
                    LookupBase = table.Column<string>(type: "text", nullable: false),
                    LookupMap = table.Column<string>(type: "text", nullable: true),
                    AuspiciousText = table.Column<string>(type: "text", nullable: true),
                    InauspiciousText = table.Column<string>(type: "text", nullable: true),
                    SpecialRule = table.Column<string>(type: "text", nullable: true),
                    SortOrder = table.Column<int>(type: "integer", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_BaziJingShenSha", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "BaziJingXiang",
                columns: table => new
                {
                    Id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    XiangType = table.Column<string>(type: "text", nullable: false),
                    Key = table.Column<string>(type: "text", nullable: false),
                    BasicImage = table.Column<string>(type: "text", nullable: true),
                    BodyImage = table.Column<string>(type: "text", nullable: true),
                    PersonImage = table.Column<string>(type: "text", nullable: true),
                    CareerImage = table.Column<string>(type: "text", nullable: true),
                    RelationImage = table.Column<string>(type: "text", nullable: true),
                    Notes = table.Column<string>(type: "text", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_BaziJingXiang", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "BaziJingYunShi",
                columns: table => new
                {
                    Id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    Category = table.Column<string>(type: "text", nullable: false),
                    Condition = table.Column<string>(type: "text", nullable: true),
                    Content = table.Column<string>(type: "text", nullable: true),
                    SortOrder = table.Column<int>(type: "integer", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_BaziJingYunShi", x => x.Id);
                });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "BaziJingCaiGuan");

            migrationBuilder.DropTable(
                name: "BaziJingConfig");

            migrationBuilder.DropTable(
                name: "BaziJingKouJue");

            migrationBuilder.DropTable(
                name: "BaziJingLiuQin");

            migrationBuilder.DropTable(
                name: "BaziJingShenSha");

            migrationBuilder.DropTable(
                name: "BaziJingXiang");

            migrationBuilder.DropTable(
                name: "BaziJingYunShi");

            migrationBuilder.DropColumn(
                name: "Id",
                schema: "public",
                table: "calendar");
        }
    }
}
