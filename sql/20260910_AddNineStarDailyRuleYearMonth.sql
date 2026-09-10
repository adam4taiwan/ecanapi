-- Migration: AddNineStarDailyRuleYearMonth
-- 2026-09-10
-- 為 NineStarDailyRules 加入流年星、流月星欄位
-- 舊資料 YearStar/MonthStar 保留為 0（不會被新查詢命中，等同自動淘汰）

ALTER TABLE "NineStarDailyRules"
    ADD COLUMN IF NOT EXISTS "YearStar"  integer NOT NULL DEFAULT 0,
    ADD COLUMN IF NOT EXISTS "MonthStar" integer NOT NULL DEFAULT 0;
