-- 2026-04-23 調整訂閱方案價格
-- BRONZE: 1200 -> 2500
-- SILVER: 1800 -> 3000
-- GOLD:   2500 -> 3600

UPDATE "MembershipPlans" SET "PriceTwd" = 2500 WHERE "Code" = 'BRONZE';
UPDATE "MembershipPlans" SET "PriceTwd" = 3000 WHERE "Code" = 'SILVER';
UPDATE "MembershipPlans" SET "PriceTwd" = 3600 WHERE "Code" = 'GOLD';
