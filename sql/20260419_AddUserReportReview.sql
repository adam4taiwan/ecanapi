-- 2026-04-19: Add review/approval fields to UserReports
-- New flow: user submits application -> admin reviews -> approves + sends email -> user downloads

ALTER TABLE "UserReports"
  ADD COLUMN IF NOT EXISTS "Status" character varying(20) NOT NULL DEFAULT 'approved',
  ADD COLUMN IF NOT EXISTS "ApprovedAt" timestamp with time zone NULL,
  ADD COLUMN IF NOT EXISTS "AdminNote" text NULL,
  ADD COLUMN IF NOT EXISTS "DownloadToken" character varying(64) NULL,
  ADD COLUMN IF NOT EXISTS "DownloadTokenExpiry" timestamp with time zone NULL;

-- Existing records are already approved (generated before this feature)
UPDATE "UserReports" SET "Status" = 'approved' WHERE "Status" = 'approved';

CREATE UNIQUE INDEX IF NOT EXISTS "IX_UserReports_DownloadToken" ON "UserReports" ("DownloadToken") WHERE "DownloadToken" IS NOT NULL;
CREATE INDEX IF NOT EXISTS "IX_UserReports_Status" ON "UserReports" ("Status");

-- EF Core migration history
INSERT INTO "__EFMigrationsHistory" ("MigrationId", "ProductVersion")
VALUES ('20260419000000_AddUserReportReview', '9.0.0')
ON CONFLICT DO NOTHING;
