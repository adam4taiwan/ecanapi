-- 命理知識庫資料表
-- 生產環境手動執行

CREATE TABLE IF NOT EXISTS "FortuneRules" (
    "Id"            SERIAL PRIMARY KEY,
    "Category"      text NOT NULL,
    "Subcategory"   text,
    "Title"         text,
    "ConditionText" text,
    "ResultText"    text NOT NULL,
    "SourceFile"    text,
    "Tags"          text,
    "IsActive"      boolean NOT NULL DEFAULT true,
    "SortOrder"     integer NOT NULL DEFAULT 0,
    "CreatedAt"     timestamp with time zone NOT NULL DEFAULT NOW(),
    "UpdatedAt"     timestamp with time zone NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS "KnowledgeDocuments" (
    "Id"             SERIAL PRIMARY KEY,
    "FileName"       text NOT NULL,
    "FileType"       text NOT NULL,
    "Category"       text,
    "ContentPreview" text,
    "RuleCount"      integer NOT NULL DEFAULT 0,
    "Status"         text NOT NULL DEFAULT 'imported',
    "UploadedAt"     timestamp with time zone NOT NULL DEFAULT NOW(),
    "UploadedBy"     text
);

-- 索引加速查詢
CREATE INDEX IF NOT EXISTS idx_fortune_rules_category ON "FortuneRules" ("Category");
CREATE INDEX IF NOT EXISTS idx_fortune_rules_active ON "FortuneRules" ("IsActive");
