using System;
using Microsoft.EntityFrameworkCore.Migrations;
using Npgsql.EntityFrameworkCore.PostgreSQL.Metadata;

#nullable disable

namespace Ecanapi.Migrations
{
    /// <inheritdoc />
    public partial class AddNineStarSystem : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "LineUsers",
                columns: table => new
                {
                    Id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    LineUserId = table.Column<string>(type: "text", nullable: false),
                    BirthYear = table.Column<int>(type: "integer", nullable: false),
                    BirthMonth = table.Column<int>(type: "integer", nullable: false),
                    BirthDay = table.Column<int>(type: "integer", nullable: false),
                    Gender = table.Column<string>(type: "text", nullable: false),
                    NatalStar = table.Column<int>(type: "integer", nullable: false),
                    DisplayName = table.Column<string>(type: "text", nullable: true),
                    CreatedAt = table.Column<DateTime>(type: "timestamp with time zone", nullable: false),
                    UpdatedAt = table.Column<DateTime>(type: "timestamp with time zone", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_LineUsers", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "NineStarDailyRules",
                columns: table => new
                {
                    Id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    NatalStar = table.Column<int>(type: "integer", nullable: false),
                    FlowStar = table.Column<int>(type: "integer", nullable: false),
                    FortuneText = table.Column<string>(type: "text", nullable: true),
                    Auspicious = table.Column<string>(type: "text", nullable: true),
                    Avoid = table.Column<string>(type: "text", nullable: true),
                    Direction = table.Column<string>(type: "text", nullable: true),
                    Color = table.Column<string>(type: "text", nullable: true),
                    UpdatedAt = table.Column<DateTime>(type: "timestamp with time zone", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_NineStarDailyRules", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "NineStarTraits",
                columns: table => new
                {
                    Id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    StarNumber = table.Column<int>(type: "integer", nullable: false),
                    StarName = table.Column<string>(type: "text", nullable: false),
                    Personality = table.Column<string>(type: "text", nullable: true),
                    Career = table.Column<string>(type: "text", nullable: true),
                    Relationship = table.Column<string>(type: "text", nullable: true),
                    Health = table.Column<string>(type: "text", nullable: true),
                    LuckyDirection = table.Column<string>(type: "text", nullable: false),
                    LuckyColor = table.Column<string>(type: "text", nullable: false),
                    LuckyNumber = table.Column<int>(type: "integer", nullable: false),
                    UpdatedAt = table.Column<DateTime>(type: "timestamp with time zone", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_NineStarTraits", x => x.Id);
                });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "LineUsers");

            migrationBuilder.DropTable(
                name: "NineStarDailyRules");

            migrationBuilder.DropTable(
                name: "NineStarTraits");
        }
    }
}
