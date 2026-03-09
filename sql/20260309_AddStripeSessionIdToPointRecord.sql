-- Migration: AddStripeSessionIdToPointRecord
-- Date: 2026-03-09
-- Purpose: Create PointRecords table (if not exists) and add Description, StripeSessionId columns.

CREATE TABLE IF NOT EXISTS "PointRecords" (
    "Id"             SERIAL PRIMARY KEY,
    "UserId"         text NOT NULL,
    "Amount"         integer NOT NULL,
    "Description"    text,
    "StripeSessionId" text,
    "CreatedAt"      timestamp with time zone NOT NULL DEFAULT NOW()
);

-- Add columns if table already existed without them
ALTER TABLE "PointRecords"
    ADD COLUMN IF NOT EXISTS "Description"     text,
    ADD COLUMN IF NOT EXISTS "StripeSessionId" text,
    ADD COLUMN IF NOT EXISTS "CreatedAt"       timestamp with time zone NOT NULL DEFAULT NOW();

-- Unique index for Stripe webhook idempotency
CREATE UNIQUE INDEX IF NOT EXISTS "IX_PointRecords_StripeSessionId"
    ON "PointRecords" ("StripeSessionId")
    WHERE "StripeSessionId" IS NOT NULL;
