-- 2026-04-19B: Add ApprovedDocxBytes for admin-corrected DOCX storage
ALTER TABLE "UserReports"
  ADD COLUMN IF NOT EXISTS "ApprovedDocxBytes" bytea NULL,
  ADD COLUMN IF NOT EXISTS "ApprovedDocxFileName" character varying(200) NULL;

INSERT INTO "__EFMigrationsHistory" ("MigrationId", "ProductVersion")
VALUES ('20260419130000_AddUserReportApprovedDocx', '9.0.0')
ON CONFLICT DO NOTHING;
