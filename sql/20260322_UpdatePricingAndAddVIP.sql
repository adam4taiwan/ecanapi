-- ============================================================
-- 20260322 Update Pricing Rules & Add VIP Plan
-- 1. Add new products: BOOK_COMBO, BOOK_LIFELONG, CONSULT_EXPLAIN
-- 2. Update plan benefits: BOOK_LIUNIAN -> BOOK_COMBO (all plans)
--    GOLD: course discount 0.8->0.9, add CONSULT_EXPLAIN quota x1
-- 3. Add VIP plan (IsActive=false until officially launched)
-- ============================================================

-- ── New products ─────────────────────────────────────────────
INSERT INTO "Products" ("Code", "Name", "Type", "PointCost", "PriceTwd", "SortOrder", "Description") VALUES
('BOOK_COMBO',      '綜合命書',   'book',         200,  NULL,  2, '八字紫微全面鑑定命書'),
('BOOK_LIFELONG',   '終身命書',   'book',         600,  NULL,  4, '終身八大運完整命書'),
('CONSULT_EXPLAIN', '玉洞子解說', 'consultation', NULL, 3600,  9, '玉洞子親自線上視訊解說，每小時 NT$3,600')
ON CONFLICT ("Code") DO NOTHING;

-- ── Update BRONZE: BOOK_LIUNIAN -> BOOK_COMBO ─────────────────
UPDATE "MembershipPlanBenefits"
SET "ProductCode" = 'BOOK_COMBO',
    "Description" = '每年綜合命書 x1'
WHERE "ProductCode" = 'BOOK_LIUNIAN'
  AND "BenefitType" = 'quota'
  AND "PlanId" = (SELECT "Id" FROM "MembershipPlans" WHERE "Code" = 'BRONZE');

-- ── Update SILVER: BOOK_LIUNIAN -> BOOK_COMBO ─────────────────
UPDATE "MembershipPlanBenefits"
SET "ProductCode" = 'BOOK_COMBO',
    "Description" = '每年綜合命書 x1'
WHERE "ProductCode" = 'BOOK_LIUNIAN'
  AND "BenefitType" = 'quota'
  AND "PlanId" = (SELECT "Id" FROM "MembershipPlans" WHERE "Code" = 'SILVER');

-- ── Update GOLD: BOOK_LIUNIAN -> BOOK_COMBO ───────────────────
UPDATE "MembershipPlanBenefits"
SET "ProductCode" = 'BOOK_COMBO',
    "Description" = '每年綜合命書 x1'
WHERE "ProductCode" = 'BOOK_LIUNIAN'
  AND "BenefitType" = 'quota'
  AND "PlanId" = (SELECT "Id" FROM "MembershipPlans" WHERE "Code" = 'GOLD');

-- ── Update GOLD: course discount 0.8 -> 0.9 ───────────────────
UPDATE "MembershipPlanBenefits"
SET "BenefitValue" = '0.9',
    "Description"  = '課程九折'
WHERE "ProductType" = 'course'
  AND "BenefitType" = 'discount'
  AND "BenefitValue" = '0.8'
  AND "PlanId" = (SELECT "Id" FROM "MembershipPlans" WHERE "Code" = 'GOLD');

-- ── Add GOLD: CONSULT_EXPLAIN quota x1 ────────────────────────
INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", 'CONSULT_EXPLAIN', NULL, 'quota', '1', '玉洞子解說 x1'
FROM "MembershipPlans" p WHERE p."Code" = 'GOLD'
AND NOT EXISTS (
  SELECT 1 FROM "MembershipPlanBenefits" b
  WHERE b."PlanId" = p."Id" AND b."ProductCode" = 'CONSULT_EXPLAIN'
);

-- ── Add VIP plan (IsActive=false) ─────────────────────────────
INSERT INTO "MembershipPlans" ("Code", "Name", "PriceTwd", "DurationDays", "IsActive", "SortOrder", "Description")
VALUES ('VIP', 'VIP 會員', 6000, 365, false, 4, '頂級尊榮方案，含終身命書與最大折扣')
ON CONFLICT ("Code") DO NOTHING;

-- VIP benefits (inserted even though IsActive=false, for when we launch)
INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", 'DAILY_FORTUNE', NULL, 'access', 'true', '每日建議存取'
FROM "MembershipPlans" p WHERE p."Code" = 'VIP'
AND NOT EXISTS (SELECT 1 FROM "MembershipPlanBenefits" b WHERE b."PlanId" = p."Id" AND b."ProductCode" = 'DAILY_FORTUNE');

INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", 'BOOK_LIFELONG', NULL, 'quota', '1', '終身命書(8大運) x1'
FROM "MembershipPlans" p WHERE p."Code" = 'VIP'
AND NOT EXISTS (SELECT 1 FROM "MembershipPlanBenefits" b WHERE b."PlanId" = p."Id" AND b."ProductCode" = 'BOOK_LIFELONG');

INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", 'CONSULT_EXPLAIN', NULL, 'quota', '1', '玉洞子解說 x1'
FROM "MembershipPlans" p WHERE p."Code" = 'VIP'
AND NOT EXISTS (SELECT 1 FROM "MembershipPlanBenefits" b WHERE b."PlanId" = p."Id" AND b."ProductCode" = 'CONSULT_EXPLAIN');

INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", 'BOOK_LIUNIAN', NULL, 'discount', '0.6', '流年命書六折優惠'
FROM "MembershipPlans" p WHERE p."Code" = 'VIP'
AND NOT EXISTS (SELECT 1 FROM "MembershipPlanBenefits" b WHERE b."PlanId" = p."Id" AND b."ProductCode" = 'BOOK_LIUNIAN' AND b."BenefitType" = 'discount');

INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", NULL, 'consultation', 'discount', '0.6', '問事六折優惠'
FROM "MembershipPlans" p WHERE p."Code" = 'VIP'
AND NOT EXISTS (SELECT 1 FROM "MembershipPlanBenefits" b WHERE b."PlanId" = p."Id" AND b."ProductType" = 'consultation' AND b."BenefitType" = 'discount');

INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", NULL, 'course', 'discount', '0.8', '課程八折優惠'
FROM "MembershipPlans" p WHERE p."Code" = 'VIP'
AND NOT EXISTS (SELECT 1 FROM "MembershipPlanBenefits" b WHERE b."PlanId" = p."Id" AND b."ProductType" = 'course' AND b."BenefitType" = 'discount');

INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", NULL, 'blessing', 'quota', '1', '贈送祈福服務 x1'
FROM "MembershipPlans" p WHERE p."Code" = 'VIP'
AND NOT EXISTS (SELECT 1 FROM "MembershipPlanBenefits" b WHERE b."PlanId" = p."Id" AND b."ProductType" = 'blessing' AND b."BenefitType" = 'quota');
