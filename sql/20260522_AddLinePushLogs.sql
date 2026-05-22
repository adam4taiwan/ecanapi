-- 新增 LINE 推播記錄表（供 AI 情境讀取）
-- 執行環境：NeonDB (production)
-- 執行日期：2026-05-22

CREATE TABLE IF NOT EXISTS "LinePushLogs" (
    "Id"           SERIAL PRIMARY KEY,
    "LineUserId"   VARCHAR(100) NOT NULL,
    "UserId"       VARCHAR(450),
    "UserEmail"    VARCHAR(256),
    "PushType"     VARCHAR(20) NOT NULL,    -- ninestar / subscriber
    "NatalStar"    INTEGER NOT NULL,
    "BirthYear"    INTEGER,
    "DayMaster"    VARCHAR(2),
    "PushDate"     DATE NOT NULL,           -- 台灣時間日期
    "Message"      TEXT NOT NULL,
    "SentAt"       TIMESTAMP NOT NULL DEFAULT NOW(),
    "Status"       VARCHAR(10) NOT NULL,    -- success / failed
    "ErrorMessage" TEXT
);

CREATE INDEX IF NOT EXISTS "IX_LinePushLogs_PushDate"    ON "LinePushLogs" ("PushDate");
CREATE INDEX IF NOT EXISTS "IX_LinePushLogs_LineUserId"  ON "LinePushLogs" ("LineUserId");
CREATE INDEX IF NOT EXISTS "IX_LinePushLogs_UserId"      ON "LinePushLogs" ("UserId");

-- 驗證
SELECT COUNT(*) FROM "LinePushLogs";
