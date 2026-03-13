-- 20260313_AddUserCharts.sql
-- Phase 3: UserCharts 資料表，儲存用戶命盤 JSON（含紫微命宮主星）

CREATE TABLE IF NOT EXISTS "UserCharts" (
    "Id" SERIAL PRIMARY KEY,
    "UserId" TEXT NOT NULL,
    "MingGongMainStars" TEXT,
    "ChartJson" TEXT NOT NULL,
    "CreatedAt" TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT NOW(),
    "UpdatedAt" TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS "IX_UserCharts_UserId" ON "UserCharts" ("UserId");

INSERT INTO "__EFMigrationsHistory" ("MigrationId", "ProductVersion")
VALUES ('20260313025235_AddUserCharts', '9.0.8')
ON CONFLICT DO NOTHING;
