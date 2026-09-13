using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Ecanapi.Migrations
{
    /// <inheritdoc />
    public partial class AddNineStarDailyRuleDayStem : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<string>(
                name: "DayStem",
                table: "NineStarDailyRules",
                type: "text",
                nullable: false,
                defaultValue: "");

            // 舊 4 維唯一索引 → 新 5 維唯一索引（加入日天干，避免同月同流日重複同一內容）
            migrationBuilder.Sql(@"DROP INDEX IF EXISTS ""idx_ninestar_daily_natal_year_month_flow"";");
            migrationBuilder.Sql(@"CREATE UNIQUE INDEX IF NOT EXISTS ""idx_ninestar_daily_natal_year_month_flow_stem""
                ON ""NineStarDailyRules"" (""NatalStar"", ""YearStar"", ""MonthStar"", ""FlowStar"", ""DayStem"");");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.Sql(@"DROP INDEX IF EXISTS ""idx_ninestar_daily_natal_year_month_flow_stem"";");
            migrationBuilder.Sql(@"CREATE UNIQUE INDEX IF NOT EXISTS ""idx_ninestar_daily_natal_year_month_flow""
                ON ""NineStarDailyRules"" (""NatalStar"", ""YearStar"", ""MonthStar"", ""FlowStar"");");

            migrationBuilder.DropColumn(
                name: "DayStem",
                table: "NineStarDailyRules");
        }
    }
}
