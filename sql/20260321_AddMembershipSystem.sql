-- ============================================================
-- 20260321 Add Membership System
-- Products / MembershipPlans / MembershipPlanBenefits /
-- UserSubscriptions / UserSubscriptionClaims
-- ============================================================

-- Products: master product catalog
CREATE TABLE IF NOT EXISTS "Products" (
    "Id"          SERIAL PRIMARY KEY,
    "Code"        VARCHAR(50)  NOT NULL UNIQUE,
    "Name"        VARCHAR(100) NOT NULL,
    "Type"        VARCHAR(30)  NOT NULL,
    "PointCost"   INTEGER,
    "PriceTwd"    INTEGER,
    "IsActive"    BOOLEAN NOT NULL DEFAULT TRUE,
    "SortOrder"   INTEGER NOT NULL DEFAULT 0,
    "Description" TEXT,
    "UpdatedAt"   TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- MembershipPlans: Gold / Silver / Bronze tier definitions
CREATE TABLE IF NOT EXISTS "MembershipPlans" (
    "Id"          SERIAL PRIMARY KEY,
    "Code"        VARCHAR(20)  NOT NULL UNIQUE,
    "Name"        VARCHAR(50)  NOT NULL,
    "PriceTwd"    INTEGER NOT NULL,
    "DurationDays" INTEGER NOT NULL DEFAULT 365,
    "IsActive"    BOOLEAN NOT NULL DEFAULT TRUE,
    "SortOrder"   INTEGER NOT NULL DEFAULT 0,
    "Description" TEXT
);

-- MembershipPlanBenefits: flexible benefit composition per plan
CREATE TABLE IF NOT EXISTS "MembershipPlanBenefits" (
    "Id"           SERIAL PRIMARY KEY,
    "PlanId"       INTEGER NOT NULL REFERENCES "MembershipPlans"("Id") ON DELETE CASCADE,
    "ProductCode"  VARCHAR(50),
    "ProductType"  VARCHAR(30),
    "BenefitType"  VARCHAR(20) NOT NULL,
    "BenefitValue" VARCHAR(20) NOT NULL,
    "Description"  TEXT
);

-- UserSubscriptions: each user's subscription record
CREATE TABLE IF NOT EXISTS "UserSubscriptions" (
    "Id"         SERIAL PRIMARY KEY,
    "UserId"     VARCHAR(450) NOT NULL,
    "PlanId"     INTEGER NOT NULL REFERENCES "MembershipPlans"("Id"),
    "StartDate"  TIMESTAMPTZ NOT NULL,
    "ExpiryDate" TIMESTAMPTZ NOT NULL,
    "Status"     VARCHAR(20) NOT NULL DEFAULT 'active',
    "PaymentRef" VARCHAR(100),
    "CreatedAt"  TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
CREATE INDEX IF NOT EXISTS "IX_UserSubscriptions_UserId" ON "UserSubscriptions"("UserId");

-- UserSubscriptionClaims: tracks free quota usage
CREATE TABLE IF NOT EXISTS "UserSubscriptionClaims" (
    "Id"             SERIAL PRIMARY KEY,
    "UserId"         VARCHAR(450) NOT NULL,
    "SubscriptionId" INTEGER NOT NULL REFERENCES "UserSubscriptions"("Id"),
    "ProductCode"    VARCHAR(50) NOT NULL,
    "ClaimYear"      INTEGER,
    "ClaimedAt"      TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- ============================================================
-- Seed: Products
-- ============================================================
INSERT INTO "Products" ("Code", "Name", "Type", "PointCost", "PriceTwd", "SortOrder") VALUES
('DAILY_FORTUNE',       '每日建議',     'daily',        NULL, NULL,  1),
('BOOK_BAZI',           '八字命書',     'book',          200, NULL,  2),
('BOOK_DAIYUN',         '大運命書',     'book',          300, NULL,  3),
('BOOK_LIUNIAN',        '流年命書',     'book',          100, NULL,  4),
('BLESSING_ANTAISUI',   '安太歲',       'blessing',     NULL, 1200,  5),
('BLESSING_LIGHT',      '光明燈',       'blessing',     NULL,  800,  6),
('BLESSING_WEALTH',     '補財庫',       'blessing',     NULL, 1500,  7),
('BLESSING_PRAYER',     '祈福服務',     'blessing',     NULL, 1000,  8),
('CONSULT_VIDEO',       '視訊問事',     'consultation', NULL, 2000,  9),
('COURSE_BASIC',        '基礎命理課程', 'course',       NULL, NULL, 10),
('LECTURE_FREE',        '不定期免費講座','lecture',      NULL, NULL, 11)
ON CONFLICT ("Code") DO NOTHING;

-- ============================================================
-- Seed: MembershipPlans (Bronze / Silver / Gold)
-- ============================================================
INSERT INTO "MembershipPlans" ("Code", "Name", "PriceTwd", "DurationDays", "SortOrder", "Description") VALUES
('BRONZE', '銅會員', 1200, 365, 1, '入門訂閱方案'),
('SILVER', '銀會員', 1800, 365, 2, '進階訂閱方案，含1項祈福服務'),
('GOLD',   '金會員', 2500, 365, 3, '尊榮訂閱方案，含1項祈福服務，最大折扣')
ON CONFLICT ("Code") DO NOTHING;

-- ============================================================
-- Seed: MembershipPlanBenefits
-- ============================================================

-- Bronze $1,200
-- daily fortune access
INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", 'DAILY_FORTUNE', NULL, 'access', 'true', '每日建議存取'
FROM "MembershipPlans" p WHERE p."Code" = 'BRONZE';
-- 1 free liunian book per year (annual quota)
INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", 'BOOK_LIUNIAN', NULL, 'quota', '1', '每年流年命書 x1'
FROM "MembershipPlans" p WHERE p."Code" = 'BRONZE';
-- 9-off discount on all books
INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", NULL, 'book', 'discount', '0.9', '命書九折'
FROM "MembershipPlans" p WHERE p."Code" = 'BRONZE';

-- Silver $1,800
INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", 'DAILY_FORTUNE', NULL, 'access', 'true', '每日建議存取'
FROM "MembershipPlans" p WHERE p."Code" = 'SILVER';
INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", 'BOOK_LIUNIAN', NULL, 'quota', '1', '每年流年命書 x1'
FROM "MembershipPlans" p WHERE p."Code" = 'SILVER';
INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", NULL, 'book', 'discount', '0.85', '命書八五折'
FROM "MembershipPlans" p WHERE p."Code" = 'SILVER';
INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", NULL, 'consultation', 'discount', '0.9', '問事九折'
FROM "MembershipPlans" p WHERE p."Code" = 'SILVER';
-- 1 free blessing item (quota, any blessing type)
INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", NULL, 'blessing', 'quota', '1', '贈送祈福服務 x1'
FROM "MembershipPlans" p WHERE p."Code" = 'SILVER';

-- Gold $2,500
INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", 'DAILY_FORTUNE', NULL, 'access', 'true', '每日建議存取'
FROM "MembershipPlans" p WHERE p."Code" = 'GOLD';
INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", 'BOOK_LIUNIAN', NULL, 'quota', '1', '每年流年命書 x1'
FROM "MembershipPlans" p WHERE p."Code" = 'GOLD';
INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", NULL, 'book', 'discount', '0.8', '命書八折'
FROM "MembershipPlans" p WHERE p."Code" = 'GOLD';
INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", NULL, 'consultation', 'discount', '0.85', '問事八五折'
FROM "MembershipPlans" p WHERE p."Code" = 'GOLD';
INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", NULL, 'course', 'discount', '0.8', '課程八折'
FROM "MembershipPlans" p WHERE p."Code" = 'GOLD';
INSERT INTO "MembershipPlanBenefits" ("PlanId", "ProductCode", "ProductType", "BenefitType", "BenefitValue", "Description")
SELECT p."Id", NULL, 'blessing', 'quota', '1', '贈送祈福服務 x1'
FROM "MembershipPlans" p WHERE p."Code" = 'GOLD';
