-- 2026-03-27 Add LineUserId to AspNetUsers for LINE Login OAuth
ALTER TABLE "AspNetUsers" ADD COLUMN IF NOT EXISTS "LineUserId" VARCHAR(64);
