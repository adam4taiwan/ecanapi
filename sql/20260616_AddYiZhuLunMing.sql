-- 一柱論命·六十甲子日柱定數
-- 建立資料表（EF Migration 會自動建，此腳本供生產環境手動執行）

CREATE TABLE IF NOT EXISTS "YiZhuLunMings" (
    "Id"               SERIAL PRIMARY KEY,
    "DayPillar"        VARCHAR(2)   NOT NULL,
    "Personality"      TEXT,
    "Poem"             TEXT,
    "MonthlyAnalysis"  TEXT,
    "VoidAnalysis"     TEXT
);

CREATE UNIQUE INDEX IF NOT EXISTS "IX_YiZhuLunMings_DayPillar"
    ON "YiZhuLunMings" ("DayPillar");

-- 初始資料在 20260616_YiZhuLunMing_Data.sql（由 Python 解析 DOCX 產生）
