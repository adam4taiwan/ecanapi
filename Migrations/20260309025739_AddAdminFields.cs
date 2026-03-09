using System;
using Microsoft.EntityFrameworkCore.Migrations;
using Npgsql.EntityFrameworkCore.PostgreSQL.Metadata;

#nullable disable

namespace Ecanapi.Migrations
{
    /// <inheritdoc />
    public partial class AddAdminFields : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.Sql(@"ALTER TABLE ""AspNetUsers"" ADD COLUMN IF NOT EXISTS ""Address""    text;");
            migrationBuilder.Sql(@"ALTER TABLE ""AspNetUsers"" ADD COLUMN IF NOT EXISTS ""Phone""      text;");
            migrationBuilder.Sql(@"ALTER TABLE ""AspNetUsers"" ADD COLUMN IF NOT EXISTS ""PostalCode"" text;");
            migrationBuilder.Sql(@"ALTER TABLE ""AspNetUsers"" ADD COLUMN IF NOT EXISTS ""TaxId""      text;");

            migrationBuilder.Sql(@"
                CREATE TABLE IF NOT EXISTS ""AtmPaymentRequests"" (
                    ""Id""           SERIAL PRIMARY KEY,
                    ""UserId""       text NOT NULL,
                    ""PackageId""    text NOT NULL,
                    ""Points""       integer NOT NULL,
                    ""PriceTwd""     integer NOT NULL,
                    ""TransferDate"" text NOT NULL,
                    ""AccountLast5"" text NOT NULL,
                    ""Status""       text NOT NULL DEFAULT 'pending',
                    ""AdminNote""    text,
                    ""CreatedAt""    timestamp with time zone NOT NULL DEFAULT NOW(),
                    ""ProcessedAt""  timestamp with time zone
                );
            ");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "AtmPaymentRequests");

            migrationBuilder.DropColumn(
                name: "Address",
                table: "AspNetUsers");

            migrationBuilder.DropColumn(
                name: "Phone",
                table: "AspNetUsers");

            migrationBuilder.DropColumn(
                name: "PostalCode",
                table: "AspNetUsers");

            migrationBuilder.DropColumn(
                name: "TaxId",
                table: "AspNetUsers");
        }
    }
}
