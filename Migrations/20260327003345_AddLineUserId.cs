using System;
using Microsoft.EntityFrameworkCore.Migrations;
using Npgsql.EntityFrameworkCore.PostgreSQL.Metadata;

#nullable disable

namespace Ecanapi.Migrations
{
    /// <inheritdoc />
    public partial class AddLineUserId : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<string>(
                name: "LineUserId",
                table: "AspNetUsers",
                type: "text",
                nullable: true);

            migrationBuilder.CreateTable(
                name: "BaziDayPillarReadings",
                columns: table => new
                {
                    Id = table.Column<int>(type: "integer", nullable: false)
                        .Annotation("Npgsql:ValueGenerationStrategy", NpgsqlValueGenerationStrategy.IdentityByDefaultColumn),
                    Sequence = table.Column<int>(type: "integer", nullable: false),
                    DayPillar = table.Column<string>(type: "character varying(2)", maxLength: 2, nullable: false),
                    DayStem = table.Column<string>(type: "character varying(1)", maxLength: 1, nullable: false),
                    DayBranch = table.Column<string>(type: "character varying(1)", maxLength: 1, nullable: false),
                    Overview = table.Column<string>(type: "text", nullable: true),
                    ShenAnalysis = table.Column<string>(type: "text", nullable: true),
                    InnerTraits = table.Column<string>(type: "text", nullable: true),
                    Career = table.Column<string>(type: "text", nullable: true),
                    Weaknesses = table.Column<string>(type: "text", nullable: true),
                    MotherInfluence = table.Column<string>(type: "text", nullable: true),
                    MonthInfluence = table.Column<string>(type: "text", nullable: true),
                    MaleChart = table.Column<string>(type: "text", nullable: true),
                    FemaleChart = table.Column<string>(type: "text", nullable: true),
                    SpecialHours = table.Column<string>(type: "text", nullable: true),
                    CreatedAt = table.Column<DateTime>(type: "timestamp with time zone", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_BaziDayPillarReadings", x => x.Id);
                });
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "BaziDayPillarReadings");

            migrationBuilder.DropColumn(
                name: "LineUserId",
                table: "AspNetUsers");
        }
    }
}
