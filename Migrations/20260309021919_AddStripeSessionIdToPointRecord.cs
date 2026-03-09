using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Ecanapi.Migrations
{
    /// <inheritdoc />
    public partial class AddStripeSessionIdToPointRecord : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.Sql(@"ALTER TABLE ""PointRecords"" ADD COLUMN IF NOT EXISTS ""Description"" text;");
            migrationBuilder.Sql(@"ALTER TABLE ""PointRecords"" ADD COLUMN IF NOT EXISTS ""StripeSessionId"" text;");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "Description",
                table: "PointRecords");

            migrationBuilder.DropColumn(
                name: "StripeSessionId",
                table: "PointRecords");
        }
    }
}
