using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Ecanapi.Migrations
{
    /// <inheritdoc />
    public partial class AddNineStarDailyRuleYearMonth : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<int>(
                name: "MonthStar",
                table: "NineStarDailyRules",
                type: "integer",
                nullable: false,
                defaultValue: 0);

            migrationBuilder.AddColumn<int>(
                name: "YearStar",
                table: "NineStarDailyRules",
                type: "integer",
                nullable: false,
                defaultValue: 0);

            // 舊 2 維唯一索引 → 新 4 維唯一索引（涵蓋流年、流月脈絡）
            migrationBuilder.Sql(@"DROP INDEX IF EXISTS ""idx_ninestar_daily_natal_flow"";");
            migrationBuilder.Sql(@"CREATE UNIQUE INDEX IF NOT EXISTS ""idx_ninestar_daily_natal_year_month_flow""
                ON ""NineStarDailyRules"" (""NatalStar"", ""YearStar"", ""MonthStar"", ""FlowStar"");");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.Sql(@"DROP INDEX IF EXISTS ""idx_ninestar_daily_natal_year_month_flow"";");
            migrationBuilder.Sql(@"CREATE UNIQUE INDEX IF NOT EXISTS ""idx_ninestar_daily_natal_flow""
                ON ""NineStarDailyRules"" (""NatalStar"", ""FlowStar"");");

            migrationBuilder.DropColumn(
                name: "MonthStar",
                table: "NineStarDailyRules");

            migrationBuilder.DropColumn(
                name: "YearStar",
                table: "NineStarDailyRules");
        }
    }
}
