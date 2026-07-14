-- 2026-07-14 Add AdviceText to BaziJingConfig and BaziJingCaiGuan
-- Purpose: C+A display approach - code generates trigger desc (C), practitioner fills advice (A)

ALTER TABLE "BaziJingConfig"
    ADD COLUMN IF NOT EXISTS "AdviceText" TEXT NULL;

ALTER TABLE "BaziJingCaiGuan"
    ADD COLUMN IF NOT EXISTS "AdviceText" TEXT NULL;

-- Record migration
INSERT INTO "__EFMigrationsHistory" ("MigrationId", "ProductVersion")
VALUES ('20260714000000_AddBaziJingAdviceText', '8.0.0')
ON CONFLICT DO NOTHING;
