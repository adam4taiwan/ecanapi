-- 2026-04-06: Add UserReports table for storing generated fortune reports
CREATE TABLE "UserReports" (
    "Id" uuid NOT NULL DEFAULT gen_random_uuid(),
    "UserId" text NOT NULL,
    "ReportType" character varying(20) NOT NULL,
    "Title" character varying(100) NOT NULL,
    "Content" text NOT NULL,
    "Parameters" jsonb NULL,
    "CreatedAt" timestamp with time zone NOT NULL DEFAULT now(),
    CONSTRAINT "PK_UserReports" PRIMARY KEY ("Id")
);

CREATE INDEX "IX_UserReports_UserId" ON "UserReports" ("UserId");
CREATE INDEX "IX_UserReports_CreatedAt" ON "UserReports" ("CreatedAt" DESC);

-- EF Core migration history
INSERT INTO "__EFMigrationsHistory" ("MigrationId", "ProductVersion")
VALUES ('20260406060408_AddUserReports', '9.0.0');
