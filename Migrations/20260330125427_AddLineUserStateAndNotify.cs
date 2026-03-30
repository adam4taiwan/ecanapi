using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Ecanapi.Migrations
{
    /// <inheritdoc />
    public partial class AddLineUserStateAndNotify : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<bool>(
                name: "NotifyEnabled",
                table: "LineUsers",
                type: "boolean",
                nullable: false,
                defaultValue: false);

            migrationBuilder.AddColumn<string>(
                name: "State",
                table: "LineUsers",
                type: "character varying(20)",
                maxLength: 20,
                nullable: false,
                defaultValue: "");

            migrationBuilder.AddColumn<int>(
                name: "TempDay",
                table: "LineUsers",
                type: "integer",
                nullable: true);

            migrationBuilder.AddColumn<int>(
                name: "TempMonth",
                table: "LineUsers",
                type: "integer",
                nullable: true);

            migrationBuilder.AddColumn<int>(
                name: "TempYear",
                table: "LineUsers",
                type: "integer",
                nullable: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "NotifyEnabled",
                table: "LineUsers");

            migrationBuilder.DropColumn(
                name: "State",
                table: "LineUsers");

            migrationBuilder.DropColumn(
                name: "TempDay",
                table: "LineUsers");

            migrationBuilder.DropColumn(
                name: "TempMonth",
                table: "LineUsers");

            migrationBuilder.DropColumn(
                name: "TempYear",
                table: "LineUsers");
        }
    }
}
