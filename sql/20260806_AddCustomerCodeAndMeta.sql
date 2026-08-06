-- Migration: AddCustomerCodeAndMeta
-- 為 Customers 表加入客戶編號、備注、建檔時間欄位
-- 執行對象：NeonDB 生產環境

ALTER TABLE "Customers"
  ADD COLUMN IF NOT EXISTS "CustomerCode" text NOT NULL DEFAULT '',
  ADD COLUMN IF NOT EXISTS "Notes" text NULL,
  ADD COLUMN IF NOT EXISTS "CreatedAt" timestamp with time zone NOT NULL DEFAULT now();

-- 為現有資料補上 CustomerCode（用建檔時間+Id）
UPDATE "Customers"
SET "CustomerCode" = TO_CHAR(COALESCE("CreatedAt", now()) AT TIME ZONE 'Asia/Taipei', 'YYYYMMDDHH24MISS') || "Id"::text
WHERE "CustomerCode" = '';

-- 加入 EF Migrations 紀錄
INSERT INTO "__EFMigrationsHistory" ("MigrationId", "ProductVersion")
VALUES ('20260806121157_AddCustomerCodeAndMeta', '8.0.0')
ON CONFLICT DO NOTHING;
