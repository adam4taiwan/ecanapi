-- Migration: AddBookingRequests
-- Date: 2026-03-22
-- Purpose: Add BookingRequests table for blessing services and consultation appointments

CREATE TABLE IF NOT EXISTS "BookingRequests" (
    "Id" SERIAL PRIMARY KEY,
    "UserId" TEXT NULL,
    "ServiceType" VARCHAR(20) NOT NULL,
    "ServiceCode" VARCHAR(50) NULL,
    "Name" VARCHAR(100) NOT NULL,
    "ContactInfo" VARCHAR(200) NULL,
    "ContactType" VARCHAR(20) NULL,
    "BirthDate" VARCHAR(20) NULL,
    "IsLunar" BOOLEAN NOT NULL DEFAULT FALSE,
    "Notes" VARCHAR(500) NULL,
    "PreferredDate" VARCHAR(50) NULL,
    "Status" VARCHAR(20) NOT NULL DEFAULT 'pending',
    "AdminNote" VARCHAR(200) NULL,
    "CreatedAt" TIMESTAMP WITH TIME ZONE NOT NULL DEFAULT NOW()
);

-- Index for admin queries
CREATE INDEX IF NOT EXISTS "IX_BookingRequests_ServiceType" ON "BookingRequests"("ServiceType");
CREATE INDEX IF NOT EXISTS "IX_BookingRequests_Status" ON "BookingRequests"("Status");
CREATE INDEX IF NOT EXISTS "IX_BookingRequests_UserId" ON "BookingRequests"("UserId");

-- Register migration in EF history
INSERT INTO "__EFMigrationsHistory" ("MigrationId", "ProductVersion")
VALUES ('20260321125210_AddBookingRequests', '9.0.8')
ON CONFLICT DO NOTHING;
