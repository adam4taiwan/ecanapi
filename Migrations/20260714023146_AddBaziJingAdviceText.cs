using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Ecanapi.Migrations
{
    /// <inheritdoc />
    public partial class AddBaziJingAdviceText : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<string>(
                name: "AdviceText",
                table: "BaziJingConfig",
                type: "text",
                nullable: true);

            migrationBuilder.AddColumn<string>(
                name: "AdviceText",
                table: "BaziJingCaiGuan",
                type: "text",
                nullable: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "AdviceText",
                table: "BaziJingConfig");

            migrationBuilder.DropColumn(
                name: "AdviceText",
                table: "BaziJingCaiGuan");
        }
    }
}
