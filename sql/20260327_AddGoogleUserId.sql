-- 2026-03-27 Add GoogleUserId to AspNetUsers
ALTER TABLE "AspNetUsers" ADD COLUMN IF NOT EXISTS "GoogleUserId" TEXT NULL;

-- Record migration
INSERT INTO "__EFMigrationsHistory" ("MigrationId", "ProductVersion")
VALUES ('20260327021724_AddGoogleUserId', '9.0.0')
ON CONFLICT DO NOTHING;
