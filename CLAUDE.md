# CLAUDE.md — Ecanapi

此檔案為 Claude Code 在此儲存庫中工作時提供指導。

## 專案概述

**Ecanapi** 是一個 ASP.NET Core 8 Web API，提供命理計算、AI 鑑定、會員認證與支付功能，部署於 Fly.io。

**技術棧：**
- ASP.NET Core 8 Web API
- Entity Framework Core 9 + PostgreSQL (Npgsql)
- ASP.NET Identity (JWT 認證)
- Google Gemini API (gemini-2.5-flash)
- Stripe 支付
- NPOI (DOCX/XLS 產出)
- Docker 部署至 Fly.io

## 專案結構

```
Controllers/
├── AuthController.cs        # 登入/註冊/密碼/個人資料
├── ConsultationController.cs # AI 命理鑑定（扣點）
├── AstrologyController.cs   # 排盤計算/匯出
├── FortuneController.cs     # 每日運勢（免費快取）
├── PaymentController.cs     # Stripe 支付
└── ...

Data/
├── ApplicationDbContext.cs  # EF Core DbContext
└── ApplicationUser.cs       # Identity 用戶（含生辰欄位）

Models/
├── DailyFortune.cs          # 每日運勢快取
└── ...

Services/
├── AstrologyService.cs      # 八字/紫微計算
├── AnalysisReportService.cs # 專家命書 DOCX 產生
└── ...

Migrations/                  # EF Core 遷移檔
sql/                         # ⭐ 生產環境手動執行的 SQL 腳本
```

## 資料庫異動規範（重要）

**每次變更資料庫結構，必須同時完成以下兩件事：**

### 1. 建立 EF Core Migration（本地開發用）
```bash
export PATH="$PATH:$HOME/.dotnet/tools"
dotnet ef migrations add <MigrationName> --context ApplicationDbContext
```
本地開發環境（`IsDevelopment()`）會在啟動時自動執行 `context.Database.Migrate()`。

### 2. 在 `sql/` 目錄新增 SQL 腳本（生產環境用）

檔名格式：`YYYYMMDD_<描述>.sql`

例如：`sql/20260308_AddDailyFortune.sql`

內容必須包含：
- 建表/修改欄位的 SQL
- EF 遷移紀錄的 INSERT：
  ```sql
  INSERT INTO "__EFMigrationsHistory" ("MigrationId", "ProductVersion")
  VALUES ('<MigrationId>', '9.0.8');
  ```

**原因：** 生產環境 (Fly.io PostgreSQL) 不自動執行 migration，需人工在 DB console 執行 SQL 腳本。

### 生成 SQL 腳本指令
```bash
export PATH="$PATH:$HOME/.dotnet/tools"
cd /home/adamtsai/projects/Ecanapi
dotnet ef migrations script <FromMigration> <ToMigration> --context ApplicationDbContext
```

## 環境設定

- `appsettings.json` — 已加入 `.gitignore`（含 API keys，不可 commit）
- Fly.io secrets 管理敏感設定：
  ```
  fly secrets set Gemini__ApiKey=... JWT__Secret=... (在 Windows 執行)
  ```
- 本地開發使用 `appsettings.json`（含 localhost PostgreSQL）

## 關鍵設定

```json
"Admin": { "Email": "adam4taiwan@gmail.com" }
```
管理員判斷用於：
- `AstrologyController.analyze` — 專家命書下載（僅管理員）
- `AuthController.profile` — `isAdmin` 欄位

## 點數費用表

| 功能 | 費用 |
|------|------|
| 每日運勢 | 免費 |
| 問事 | 10 點 |
| 綜合鑑定 | 50 點 |
| 大運 5 年 | 150 點 |
| 大運 10 年 | 200 點 |
| 大運 20 年 | 250 點 |
| 大運 30 年 | 300 點 |
| 大運終身 | 500 點 |
| 流年命書 | 20 點 |
| 專家命書(管理員) | 200 點 |

## 注意事項

- Gemini API 呼叫 timeout 設為 5 分鐘（ConsultationController）、3 分鐘（FortuneController）
- 使用 `ILogger` 而非 `Console.WriteLine`（Fly.io logs 只顯示 ILogger 輸出）
- CORS 允許 `http://localhost:3000` 和 `https://myweb.fly.dev`
- 部署指令在 Windows 執行：`fly deploy --local-only`
