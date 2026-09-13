-- 九星每日規則 KB 加入日天干欄位
-- 目的：同月同流日星可能重複推同樣內容（每9日循環），加入日干（甲~癸）後每日組合唯一
-- 舊資料 DayStem="" 視為舊版，不影響新規則查詢

ALTER TABLE "NineStarDailyRules"
    ADD COLUMN IF NOT EXISTS "DayStem" text NOT NULL DEFAULT '';

-- 更新唯一索引：從4維升級為5維（NatalStar, YearStar, MonthStar, FlowStar, DayStem）
DROP INDEX IF EXISTS "idx_ninestar_daily_natal_year_month_flow";
CREATE UNIQUE INDEX IF NOT EXISTS "idx_ninestar_daily_natal_year_month_flow_stem"
    ON "NineStarDailyRules" ("NatalStar", "YearStar", "MonthStar", "FlowStar", "DayStem");
