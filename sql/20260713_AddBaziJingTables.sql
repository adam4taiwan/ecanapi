-- 八字真經知識庫 - 建表（生產環境手動執行）
-- 對應 Migration: AddBaziJingKnowledgeBase

CREATE TABLE IF NOT EXISTS "BaziJingConfig" (
    "Id" SERIAL PRIMARY KEY,
    "ConfigType" VARCHAR(10) NOT NULL,   -- 吉 or 凶
    "ConfigName" VARCHAR(50) NOT NULL,
    "Condition" VARCHAR(200),
    "Content" TEXT,
    "SortOrder" INT NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS "BaziJingCaiGuan" (
    "Id" SERIAL PRIMARY KEY,
    "Category" VARCHAR(20) NOT NULL,     -- 財 / 官 / 互動
    "ConfigType" VARCHAR(50) NOT NULL,
    "Condition" VARCHAR(200),
    "Content" TEXT,
    "SortOrder" INT NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS "BaziJingXiang" (
    "Id" SERIAL PRIMARY KEY,
    "XiangType" VARCHAR(5) NOT NULL,     -- 天干 or 地支
    "Key" VARCHAR(5) NOT NULL,
    "BasicImage" TEXT,
    "BodyImage" TEXT,
    "PersonImage" TEXT,
    "CareerImage" TEXT,
    "RelationImage" TEXT,
    "Notes" TEXT
);

CREATE TABLE IF NOT EXISTS "BaziJingShenSha" (
    "Id" SERIAL PRIMARY KEY,
    "Name" VARCHAR(20) NOT NULL,
    "LookupBase" VARCHAR(10) NOT NULL,   -- 年支 / 日干 / 日支
    "LookupMap" VARCHAR(200),            -- JSON map
    "AuspiciousText" TEXT,
    "InauspiciousText" TEXT,
    "SpecialRule" TEXT,
    "SortOrder" INT NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS "BaziJingKouJue" (
    "Id" SERIAL PRIMARY KEY,
    "Category" VARCHAR(20) NOT NULL,     -- 年柱/月柱/日柱/時柱/十排歌/十神口訣/兩神組合
    "Condition" VARCHAR(100),
    "Content" TEXT,
    "SortOrder" INT NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS "BaziJingLiuQin" (
    "Id" SERIAL PRIMARY KEY,
    "LiuQinType" VARCHAR(10) NOT NULL,   -- 父/母/兄弟/配偶/子女
    "Category" VARCHAR(30) NOT NULL,     -- 個數/時機/品質/克損
    "Condition" VARCHAR(200),
    "Content" TEXT,
    "SortOrder" INT NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS "BaziJingYunShi" (
    "Id" SERIAL PRIMARY KEY,
    "Category" VARCHAR(20) NOT NULL,     -- 大運/流年/共通
    "Title" VARCHAR(50),                 -- 短標題 e.g. 喜用運, 忌神運
    "Condition" VARCHAR(200),            -- 詳細條件描述
    "Content" TEXT,
    "SortOrder" INT NOT NULL DEFAULT 0
);

-- Migration history
INSERT INTO "__EFMigrationsHistory" ("MigrationId", "ProductVersion")
VALUES ('20260713000000_AddBaziJingKnowledgeBase', '8.0.0')
ON CONFLICT DO NOTHING;
