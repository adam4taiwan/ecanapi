using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Ecanapi.Migrations
{
    /// <inheritdoc />
    public partial class EnhanceBaziDailyPush : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<int>(
                name: "BodyPct",
                table: "UserCharts",
                type: "integer",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "JiShenElem",
                table: "UserCharts",
                type: "character varying(2)",
                maxLength: 2,
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "Pattern",
                table: "UserCharts",
                type: "character varying(20)",
                maxLength: 20,
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "YongShenElem",
                table: "UserCharts",
                type: "character varying(2)",
                maxLength: 2,
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "ContentCategory",
                table: "LinePushLogs",
                type: "character varying(20)",
                maxLength: 20,
                nullable: true);

            migrationBuilder.AddColumn<bool>(
                name: "IsShun",
                table: "LinePushLogs",
                type: "boolean",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "ShenShaHit",
                table: "LinePushLogs",
                type: "character varying(50)",
                maxLength: 50,
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "ShiShen",
                table: "LinePushLogs",
                type: "character varying(6)",
                maxLength: 6,
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "TodayGanZhi",
                table: "LinePushLogs",
                type: "character varying(4)",
                maxLength: 4,
                nullable: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "BodyPct",
                table: "UserCharts");

            migrationBuilder.DropColumn(
                name: "JiShenElem",
                table: "UserCharts");

            migrationBuilder.DropColumn(
                name: "Pattern",
                table: "UserCharts");

            migrationBuilder.DropColumn(
                name: "YongShenElem",
                table: "UserCharts");

            migrationBuilder.DropColumn(
                name: "ContentCategory",
                table: "LinePushLogs");

            migrationBuilder.DropColumn(
                name: "IsShun",
                table: "LinePushLogs");

            migrationBuilder.DropColumn(
                name: "ShenShaHit",
                table: "LinePushLogs");

            migrationBuilder.DropColumn(
                name: "ShiShen",
                table: "LinePushLogs");

            migrationBuilder.DropColumn(
                name: "TodayGanZhi",
                table: "LinePushLogs");
        }
    }
}
