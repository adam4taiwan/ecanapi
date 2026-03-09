-- Migration: AddAdminFields
-- Date: 2026-03-09
-- Purpose:
--   1. Add commercial fields (Phone, PostalCode, Address, TaxId) to AspNetUsers
--   2. Create AtmPaymentRequests table for manual ATM transfer review

-- 1. Add user commercial fields
ALTER TABLE "AspNetUsers"
    ADD COLUMN IF NOT EXISTS "Phone"      text,
    ADD COLUMN IF NOT EXISTS "PostalCode" text,
    ADD COLUMN IF NOT EXISTS "Address"    text,
    ADD COLUMN IF NOT EXISTS "TaxId"      text;

-- 2. Create ATM payment request table
CREATE TABLE IF NOT EXISTS "AtmPaymentRequests" (
    "Id"           SERIAL PRIMARY KEY,
    "UserId"       text NOT NULL,
    "PackageId"    text NOT NULL,
    "Points"       integer NOT NULL,
    "PriceTwd"     integer NOT NULL,
    "TransferDate" text NOT NULL,
    "AccountLast5" text NOT NULL,
    "Status"       text NOT NULL DEFAULT 'pending',
    "AdminNote"    text,
    "CreatedAt"    timestamp with time zone NOT NULL DEFAULT NOW(),
    "ProcessedAt"  timestamp with time zone
);

CREATE INDEX IF NOT EXISTS "IX_AtmPaymentRequests_Status"
    ON "AtmPaymentRequests" ("Status");

CREATE INDEX IF NOT EXISTS "IX_AtmPaymentRequests_UserId"
    ON "AtmPaymentRequests" ("UserId");
