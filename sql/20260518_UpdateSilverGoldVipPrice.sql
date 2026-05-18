-- 2026-05-18 更新銀/金/VIP 會員方案價格
UPDATE "MembershipPlans" SET "PriceTwd" = 4800 WHERE "Code" = 'SILVER';
UPDATE "MembershipPlans" SET "PriceTwd" = 6000 WHERE "Code" = 'GOLD';
UPDATE "MembershipPlans" SET "PriceTwd" = 8000 WHERE "Code" = 'VIP';
