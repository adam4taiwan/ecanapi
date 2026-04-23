-- 2026-04-23 新增 VIP 會員方案
-- VIP $6,000/年：玉洞子傳家寶典 x1 + 流年命書 x1（不含大運命書，已內含8大運）

-- ============================================================
-- 1. Products: 新增 BOOK_VIP 商品
-- ============================================================
INSERT INTO "Products" ("Code", "Name", "ProductType", "PriceTwd", "Description", "IsActive")
VALUES ('BOOK_VIP', '玉洞子傳家寶典', 'book', 6000, '17章完整玉洞子命書，含8大運完整論述', true)
ON CONFLICT ("Code") DO NOTHING;

-- ============================================================
-- 2. MembershipPlans: 新增 VIP 方案
-- ============================================================
INSERT INTO "MembershipPlans" ("Code", "Name", "PriceTwd", "DurationDays", "SortOrder", "Description")
VALUES ('VIP', 'VIP會員', 6000, 365, 4, '頂級尊榮，玉洞子完整命書含8大運，每年流年命書')
ON CONFLICT ("Code") DO NOTHING;

-- ============================================================
-- 3. MembershipPlanBenefits: VIP 福利設定
-- ============================================================

-- 玉洞子傳家寶典 1次/年（BOOK_VIP）
INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", 'BOOK_VIP', NULL, 'quota', '1', '玉洞子傳家寶典 1次/年'
FROM "MembershipPlans" p WHERE p."Code" = 'VIP'
  AND NOT EXISTS (
    SELECT 1 FROM "MembershipPlanBenefits" b JOIN "MembershipPlans" p2 ON p2."Id" = b."PlanId"
    WHERE p2."Code" = 'VIP' AND b."ProductCode" = 'BOOK_VIP'
  );

-- 流年命書 1次/年
INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", 'BOOK_LIUNIAN', NULL, 'quota', '1', '流年命書 1次/年'
FROM "MembershipPlans" p WHERE p."Code" = 'VIP'
  AND NOT EXISTS (
    SELECT 1 FROM "MembershipPlanBenefits" b JOIN "MembershipPlans" p2 ON p2."Id" = b."PlanId"
    WHERE p2."Code" = 'VIP' AND b."ProductCode" = 'BOOK_LIUNIAN'
  );

-- 每日建議無限存取
INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", 'DAILY_FORTUNE', NULL, 'access', 'true', '每日建議存取'
FROM "MembershipPlans" p WHERE p."Code" = 'VIP'
  AND NOT EXISTS (
    SELECT 1 FROM "MembershipPlanBenefits" b JOIN "MembershipPlans" p2 ON p2."Id" = b."PlanId"
    WHERE p2."Code" = 'VIP' AND b."ProductCode" = 'DAILY_FORTUNE'
  );
