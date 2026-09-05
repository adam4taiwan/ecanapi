-- 強化八字流日推播：UserCharts 新增用神欄位 + LinePushLogs 新增分類欄位
-- 執行環境：NeonDB (production)
-- 執行日期：2026-09-05

-- UserCharts：新增用神/忌神/格局/身強弱欄位（供每日推播喜忌判定）
ALTER TABLE "UserCharts"
    ADD COLUMN IF NOT EXISTS "YongShenElem" VARCHAR(2),
    ADD COLUMN IF NOT EXISTS "JiShenElem"   VARCHAR(2),
    ADD COLUMN IF NOT EXISTS "Pattern"      VARCHAR(20),
    ADD COLUMN IF NOT EXISTS "BodyPct"      INTEGER;

-- LinePushLogs：新增內容分類與八字流日元數據欄位
ALTER TABLE "LinePushLogs"
    ADD COLUMN IF NOT EXISTS "ContentCategory" VARCHAR(20),
    ADD COLUMN IF NOT EXISTS "TodayGanZhi"     VARCHAR(4),
    ADD COLUMN IF NOT EXISTS "ShiShen"         VARCHAR(6),
    ADD COLUMN IF NOT EXISTS "IsShun"          BOOLEAN,
    ADD COLUMN IF NOT EXISTS "ShenShaHit"      VARCHAR(50);

-- 建立索引（供分析查詢用）
CREATE INDEX IF NOT EXISTS "IX_LinePushLogs_ContentCategory" ON "LinePushLogs" ("ContentCategory");
CREATE INDEX IF NOT EXISTS "IX_LinePushLogs_ShenShaHit"      ON "LinePushLogs" ("ShenShaHit") WHERE "ShenShaHit" IS NOT NULL;

-- 驗證
SELECT column_name, data_type
FROM information_schema.columns
WHERE table_name = 'UserCharts'
  AND column_name IN ('YongShenElem','JiShenElem','Pattern','BodyPct')
ORDER BY column_name;

SELECT column_name, data_type
FROM information_schema.columns
WHERE table_name = 'LinePushLogs'
  AND column_name IN ('ContentCategory','TodayGanZhi','ShiShen','IsShun','ShenShaHit')
ORDER BY column_name;
