-- Patch: BaziJingYunShi 新增 Title 欄位
-- 若表已建立但缺少 Title 欄位，執行此腳本

ALTER TABLE "BaziJingYunShi" ADD COLUMN IF NOT EXISTS "Title" VARCHAR(50);
