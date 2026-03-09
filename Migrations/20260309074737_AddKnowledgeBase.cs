using System;
using Microsoft.EntityFrameworkCore.Migrations;
using Npgsql.EntityFrameworkCore.PostgreSQL.Metadata;

#nullable disable

namespace Ecanapi.Migrations
{
    /// <inheritdoc />
    public partial class AddKnowledgeBase : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.Sql(@"
                CREATE TABLE IF NOT EXISTS ""FortuneRules"" (
                    ""Id""            SERIAL PRIMARY KEY,
                    ""Category""      text NOT NULL,
                    ""Subcategory""   text,
                    ""Title""         text,
                    ""ConditionText"" text,
                    ""ResultText""    text NOT NULL,
                    ""SourceFile""    text,
                    ""Tags""          text,
                    ""IsActive""      boolean NOT NULL DEFAULT true,
                    ""SortOrder""     integer NOT NULL DEFAULT 0,
                    ""CreatedAt""     timestamp with time zone NOT NULL DEFAULT NOW(),
                    ""UpdatedAt""     timestamp with time zone NOT NULL DEFAULT NOW()
                );
            ");

            migrationBuilder.Sql(@"
                CREATE TABLE IF NOT EXISTS ""KnowledgeDocuments"" (
                    ""Id""             SERIAL PRIMARY KEY,
                    ""FileName""       text NOT NULL,
                    ""FileType""       text NOT NULL,
                    ""Category""       text,
                    ""ContentPreview"" text,
                    ""RuleCount""      integer NOT NULL DEFAULT 0,
                    ""Status""         text NOT NULL DEFAULT 'imported',
                    ""UploadedAt""     timestamp with time zone NOT NULL DEFAULT NOW(),
                    ""UploadedBy""     text
                );
            ");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "FortuneRules");

            migrationBuilder.DropTable(
                name: "KnowledgeDocuments");
        }
    }
}
