-- 20260425: 新增學生白名單資料表
-- 供系統管理員設定允許使用學生練習排盤功能的登入者 email

CREATE TABLE IF NOT EXISTS "StudentWhiteLists" (
    "Id" SERIAL PRIMARY KEY,
    "Email" VARCHAR(255) NOT NULL,
    "Note" VARCHAR(255) NULL,
    "AddedByEmail" VARCHAR(255) NOT NULL DEFAULT '',
    "AddedAt" TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE UNIQUE INDEX IF NOT EXISTS "IX_StudentWhiteLists_Email"
    ON "StudentWhiteLists" ("Email");

-- 記錄 EF Migration
INSERT INTO "__EFMigrationsHistory" ("MigrationId", "ProductVersion")
VALUES ('20260425140413_AddStudentWhiteList', '8.0.0')
ON CONFLICT DO NOTHING;
