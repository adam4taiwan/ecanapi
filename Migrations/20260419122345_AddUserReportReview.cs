using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Ecanapi.Migrations
{
    /// <inheritdoc />
    public partial class AddUserReportReview : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<string>(
                name: "AdminNote",
                table: "UserReports",
                type: "text",
                nullable: true);

            migrationBuilder.AddColumn<DateTime>(
                name: "ApprovedAt",
                table: "UserReports",
                type: "timestamp with time zone",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "DownloadToken",
                table: "UserReports",
                type: "character varying(64)",
                maxLength: 64,
                nullable: true);

            migrationBuilder.AddColumn<DateTime>(
                name: "DownloadTokenExpiry",
                table: "UserReports",
                type: "timestamp with time zone",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "Status",
                table: "UserReports",
                type: "character varying(20)",
                maxLength: 20,
                nullable: false,
                defaultValue: "");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "AdminNote",
                table: "UserReports");

            migrationBuilder.DropColumn(
                name: "ApprovedAt",
                table: "UserReports");

            migrationBuilder.DropColumn(
                name: "DownloadToken",
                table: "UserReports");

            migrationBuilder.DropColumn(
                name: "DownloadTokenExpiry",
                table: "UserReports");

            migrationBuilder.DropColumn(
                name: "Status",
                table: "UserReports");
        }
    }
}
