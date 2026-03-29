-- ============================================================
-- 20260329 Update Membership Subscription Model
-- New pricing: BRONZE 1000 / SILVER 2000 / GOLD 3000
-- New benefits: quota-based per book type, no points
-- ============================================================

-- Update plan prices
UPDATE "MembershipPlans" SET "PriceTwd" = 1000 WHERE "Code" = 'BRONZE';
UPDATE "MembershipPlans" SET "PriceTwd" = 2000 WHERE "Code" = 'SILVER';
UPDATE "MembershipPlans" SET "PriceTwd" = 3000 WHERE "Code" = 'GOLD';

-- Clear all existing benefits
DELETE FROM "MembershipPlanBenefits";

-- ============================================================
-- BRONZE NT$1,000: 每日運勢(無限) + 八字命書 x1/年
-- ============================================================
INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", 'DAILY_FORTUNE', NULL, 'access', 'true', '每日建議無限存取'
FROM "MembershipPlans" p WHERE p."Code" = 'BRONZE';

INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", 'BOOK_BAZI', NULL, 'quota', '1', '八字命書 1次/年'
FROM "MembershipPlans" p WHERE p."Code" = 'BRONZE';

-- ============================================================
-- SILVER NT$2,000: 每日運勢(無限) + 八字命書 x1/年 + 流年命書 x1/年 + 問事可用
-- ============================================================
INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", 'DAILY_FORTUNE', NULL, 'access', 'true', '每日建議無限存取'
FROM "MembershipPlans" p WHERE p."Code" = 'SILVER';

INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", 'BOOK_BAZI', NULL, 'quota', '1', '八字命書 1次/年'
FROM "MembershipPlans" p WHERE p."Code" = 'SILVER';

INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", 'BOOK_LIUNIAN', NULL, 'quota', '1', '流年命書 1次/年'
FROM "MembershipPlans" p WHERE p."Code" = 'SILVER';

INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", 'TOPIC_CONSULT', NULL, 'access', 'true', '問事命書可用'
FROM "MembershipPlans" p WHERE p."Code" = 'SILVER';

-- ============================================================
-- GOLD NT$3,000: 每日運勢(無限) + 八字命書 x1/年 + 流年命書 x1/年 + 大運命書 x1/年 + 問事可用
-- ============================================================
INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", 'DAILY_FORTUNE', NULL, 'access', 'true', '每日建議無限存取'
FROM "MembershipPlans" p WHERE p."Code" = 'GOLD';

INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", 'BOOK_BAZI', NULL, 'quota', '1', '八字命書 1次/年'
FROM "MembershipPlans" p WHERE p."Code" = 'GOLD';

INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", 'BOOK_LIUNIAN', NULL, 'quota', '1', '流年命書 1次/年'
FROM "MembershipPlans" p WHERE p."Code" = 'GOLD';

INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", 'BOOK_DAIYUN', NULL, 'quota', '1', '大運命書 1次/年'
FROM "MembershipPlans" p WHERE p."Code" = 'GOLD';

INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", 'TOPIC_CONSULT', NULL, 'access', 'true', '問事命書可用'
FROM "MembershipPlans" p WHERE p."Code" = 'GOLD';
