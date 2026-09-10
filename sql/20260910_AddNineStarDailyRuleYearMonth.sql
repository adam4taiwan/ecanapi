-- Migration: AddNineStarDailyRuleYearMonth
-- 2026-09-10
-- 為 NineStarDailyRules 加入流年星、流月星欄位
-- 舊資料 YearStar/MonthStar 保留為 0（不會被新查詢命中，等同自動淘汰）

ALTER TABLE "NineStarDailyRules"
    ADD COLUMN IF NOT EXISTS "YearStar"  integer NOT NULL DEFAULT 0,
    ADD COLUMN IF NOT EXISTS "MonthStar" integer NOT NULL DEFAULT 0;

-- 重建唯一索引：舊 2 維 (NatalStar, FlowStar) → 新 4 維 (NatalStar, YearStar, MonthStar, FlowStar)
DROP INDEX IF EXISTS "idx_ninestar_daily_natal_flow";
CREATE UNIQUE INDEX IF NOT EXISTS "idx_ninestar_daily_natal_year_month_flow"
    ON "NineStarDailyRules" ("NatalStar", "YearStar", "MonthStar", "FlowStar");
