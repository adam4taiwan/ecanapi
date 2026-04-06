-- 2026-04-06: Add TrialStartDate to AspNetUsers for 7-day free trial feature
ALTER TABLE "AspNetUsers" ADD COLUMN "TrialStartDate" timestamp with time zone NULL;

-- EF Core migration history
INSERT INTO "__EFMigrationsHistory" ("MigrationId", "ProductVersion")
VALUES ('20260406060013_AddTrialStartDate', '9.0.0');
