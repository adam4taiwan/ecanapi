-- 九星氣學系統：NineStarTraits / NineStarDailyRules / LineUsers
-- 執行環境：NeonDB 生產環境（手動執行）
-- 日期：2026-03-27

-- 1. 九星本命星特質 KB
CREATE TABLE IF NOT EXISTS "NineStarTraits" (
    "Id"             SERIAL PRIMARY KEY,
    "StarNumber"     INTEGER NOT NULL,
    "StarName"       TEXT NOT NULL,
    "Personality"    TEXT,
    "Career"         TEXT,
    "Relationship"   TEXT,
    "Health"         TEXT,
    "LuckyDirection" TEXT NOT NULL DEFAULT '',
    "LuckyColor"     TEXT NOT NULL DEFAULT '',
    "LuckyNumber"    INTEGER NOT NULL DEFAULT 0,
    "UpdatedAt"      TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT NOW()
);

CREATE UNIQUE INDEX IF NOT EXISTS idx_ninestar_traits_number
    ON "NineStarTraits" ("StarNumber");

-- 2. 九星每日建議 KB（本命星 × 流星 81 組合）
CREATE TABLE IF NOT EXISTS "NineStarDailyRules" (
    "Id"          SERIAL PRIMARY KEY,
    "NatalStar"   INTEGER NOT NULL,
    "FlowStar"    INTEGER NOT NULL,
    "FortuneText" TEXT,
    "Auspicious"  TEXT,
    "Avoid"       TEXT,
    "Direction"   TEXT,
    "Color"       TEXT,
    "UpdatedAt"   TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT NOW()
);

CREATE UNIQUE INDEX IF NOT EXISTS idx_ninestar_daily_natal_flow
    ON "NineStarDailyRules" ("NatalStar", "FlowStar");

-- 3. LINE Bot 用戶生辰資料
CREATE TABLE IF NOT EXISTS "LineUsers" (
    "Id"          SERIAL PRIMARY KEY,
    "LineUserId"  TEXT NOT NULL,
    "BirthYear"   INTEGER NOT NULL,
    "BirthMonth"  INTEGER NOT NULL,
    "BirthDay"    INTEGER NOT NULL,
    "Gender"      TEXT NOT NULL,
    "NatalStar"   INTEGER NOT NULL,
    "DisplayName" TEXT,
    "CreatedAt"   TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT NOW(),
    "UpdatedAt"   TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT NOW()
);

CREATE UNIQUE INDEX IF NOT EXISTS idx_lineusers_lineuserid
    ON "LineUsers" ("LineUserId");

-- 4. 預填九星基本資料（方位、顏色、數字）
-- Gemini 會自動補充 Personality/Career/Relationship/Health
INSERT INTO "NineStarTraits" ("StarNumber","StarName","LuckyDirection","LuckyColor","LuckyNumber","UpdatedAt")
VALUES
    (1,'一白水星','北方','白色、藍色',1,NOW()),
    (2,'二黑土星','西南方','黃色、棕色',2,NOW()),
    (3,'三碧木星','東方','綠色、青色',3,NOW()),
    (4,'四綠木星','東南方','綠色、青色',4,NOW()),
    (5,'五黃土星','中宮（避中央）','黃色（需化解）',5,NOW()),
    (6,'六白金星','西北方','白色、金色',6,NOW()),
    (7,'七赤金星','西方','白色、金色',7,NOW()),
    (8,'八白土星','東北方','白色、黃色',8,NOW()),
    (9,'九紫火星','南方','紫色、紅色',9,NOW())
ON CONFLICT ("StarNumber") DO NOTHING;

-- 5. 記錄 Migration 歷史（避免 EF 自動執行時重複）
INSERT INTO "__EFMigrationsHistory" ("MigrationId","ProductVersion")
VALUES ('20260327000000_AddNineStarSystem','8.0.0')
ON CONFLICT DO NOTHING;
