-- 修正四個方案定價 + 啟用 VIP + 新增 VIP benefits
-- 執行環境：NeonDB (production)
-- 執行日期：2026-05-21

-- 1. 修正定價
UPDATE "MembershipPlans" SET "PriceTwd" = 3600 WHERE "Code" = 'BRONZE';
UPDATE "MembershipPlans" SET "PriceTwd" = 4800 WHERE "Code" = 'SILVER';
UPDATE "MembershipPlans" SET "PriceTwd" = 6000 WHERE "Code" = 'GOLD';
UPDATE "MembershipPlans" SET "PriceTwd" = 8000, "IsActive" = true WHERE "Code" = 'VIP';

-- 2. 新增 VIP benefits（若已存在先清除再插入）
DELETE FROM "MembershipPlanBenefits"
WHERE "PlanId" = (SELECT "Id" FROM "MembershipPlans" WHERE "Code" = 'VIP');

INSERT INTO "MembershipPlanBenefits" ("PlanId", "BenefitType", "ProductCode", "ProductType", "BenefitValue", "Description")
SELECT p."Id", b."BenefitType", b."ProductCode", b."ProductType", b."BenefitValue", b."Description"
FROM (VALUES
  ('access', 'DAILY_FORTUNE', NULL, 'true', '每日建議無限存取'),
  ('access', 'TOPIC_CONSULT', NULL, 'true', '問事命書可用'),
  ('quota',  'BOOK_VIP',      NULL, '1',    '玉洞子傳家寶典 1次/年'),
  ('quota',  'BOOK_LIUNIAN',  NULL, '1',    '流年命書 1次/年')
) AS b("BenefitType", "ProductCode", "ProductType", "BenefitValue", "Description")
CROSS JOIN (SELECT "Id" FROM "MembershipPlans" WHERE "Code" = 'VIP') AS p;

-- 3. 驗證
SELECT "Code", "Name", "PriceTwd", "IsActive" FROM "MembershipPlans" ORDER BY "SortOrder";

SELECT p."Code", b."BenefitType", b."ProductCode", b."BenefitValue", b."Description"
FROM "MembershipPlanBenefits" b
JOIN "MembershipPlans" p ON b."PlanId" = p."Id"
WHERE p."Code" IN ('GOLD', 'VIP')
ORDER BY p."SortOrder", b."BenefitType", b."ProductCode";
