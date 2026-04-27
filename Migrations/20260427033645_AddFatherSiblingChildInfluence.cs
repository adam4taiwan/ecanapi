using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Ecanapi.Migrations
{
    /// <inheritdoc />
    public partial class AddFatherSiblingChildInfluence : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<string>(
                name: "ChildInfluence",
                table: "BaziDayPillarReadings",
                type: "text",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "FatherInfluence",
                table: "BaziDayPillarReadings",
                type: "text",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "SiblingInfluence",
                table: "BaziDayPillarReadings",
                type: "text",
                nullable: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "ChildInfluence",
                table: "BaziDayPillarReadings");

            migrationBuilder.DropColumn(
                name: "FatherInfluence",
                table: "BaziDayPillarReadings");

            migrationBuilder.DropColumn(
                name: "SiblingInfluence",
                table: "BaziDayPillarReadings");
        }
    }
}
