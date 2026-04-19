using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Ecanapi.Migrations
{
    /// <inheritdoc />
    public partial class AddUserReportApprovedDocx : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<byte[]>(
                name: "ApprovedDocxBytes",
                table: "UserReports",
                type: "bytea",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "ApprovedDocxFileName",
                table: "UserReports",
                type: "character varying(200)",
                maxLength: 200,
                nullable: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "ApprovedDocxBytes",
                table: "UserReports");

            migrationBuilder.DropColumn(
                name: "ApprovedDocxFileName",
                table: "UserReports");
        }
    }
}
