-- 2026-04-15: Add/update BOOK_BAZI quota for BRONZE plan
-- 玉洞子八字紫微命書改為銅會員基本福利，每年 1 次免費
-- 若已有 BOOK_BAZI quota 記錄則更新描述，否則新增

-- 更新現有記錄的描述
UPDATE "MembershipPlanBenefits"
SET "Description" = '玉洞子八字紫微命書 1次/年'
WHERE "ProductCode" = 'BOOK_BAZI'
  AND "BenefitType" = 'quota'
  AND "PlanId" = (SELECT "Id" FROM "MembershipPlans" WHERE "Code" = 'BRONZE');

-- 若無記錄則新增（UPDATE 影響 0 rows 時執行）
INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", 'BOOK_BAZI', NULL, 'quota', '1', '玉洞子八字紫微命書 1次/年'
FROM "MembershipPlans" p WHERE p."Code" = 'BRONZE'
  AND NOT EXISTS (
    SELECT 1 FROM "MembershipPlanBenefits" b2
    JOIN "MembershipPlans" p2 ON p2."Id" = b2."PlanId"
    WHERE p2."Code" = 'BRONZE' AND b2."ProductCode" = 'BOOK_BAZI' AND b2."BenefitType" = 'quota'
  );
