-- 新增父親/兄弟/子女 日柱影響欄位
ALTER TABLE "BaziDayPillarReadings" ADD COLUMN IF NOT EXISTS "FatherInfluence" TEXT;
ALTER TABLE "BaziDayPillarReadings" ADD COLUMN IF NOT EXISTS "SiblingInfluence" TEXT;
ALTER TABLE "BaziDayPillarReadings" ADD COLUMN IF NOT EXISTS "ChildInfluence" TEXT;
-- 欄位新增後可逐步填入專家內容（NULL 時程式自動依五行百分比計算）
