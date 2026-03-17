# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 專案概述

**Ecanapi** 是一個 ASP.NET Core 8 Web API，提供命理計算、AI 鑑定、會員認證與支付功能，部署於 Fly.io。

**技術棧：**
- ASP.NET Core 8 Web API
- Entity Framework Core 9 + PostgreSQL (Npgsql)
- ASP.NET Identity (JWT 認證)
- Google Gemini API (gemini-2.5-flash)
- Stripe 支付
- NPOI (DOCX/XLS 產出)
- DocumentFormat.OpenXml (DOCX 解析)
- YiJingFramework.Nongli (農曆計算)
- Docker 部署至 Fly.io

## 本地開發環境

- **本地 API 位址：** `http://localhost:5013`
- **本地資料庫：** `appsettings.json` 中的 PostgreSQL 連線字串（localhost）
- **前端 MyWeb：** `http://localhost:3000`（呼叫 `http://localhost:5013/api/...`）

## 常用指令

```bash
# 建置
export PATH="$PATH:$HOME/.dotnet/tools"
dotnet build

# 本地執行（在 port 5013）
dotnet run

# EF Core migration
dotnet ef migrations add <MigrationName> --context ApplicationDbContext

# 生成 SQL 腳本（供生產環境手動執行）
dotnet ef migrations script <FromMigration> <ToMigration> --context ApplicationDbContext
```

> 部署至 Fly.io 需在 Windows 執行：`fly deploy --local-only`

## 專案結構

```
Controllers/
├── AuthController.cs        # 登入/註冊/密碼/個人資料
├── ConsultationController.cs # AI 命理鑑定（扣點，含 Gemini 呼叫）
├── AstrologyController.cs   # 排盤計算/匯出（含專家命書，僅管理員）
├── FortuneController.cs     # 每日運勢（免費快取，含個人化版本）
├── PaymentController.cs     # Stripe 支付 + ATM 付款
├── AdminController.cs       # 管理員功能（用戶管理、點數調整）
├── KnowledgeController.cs   # 命理知識庫 CRUD（僅管理員）
├── AnalysisController.cs    # 紫微斗數分析
├── CalculationController.cs # 八字/排盤計算
├── CalendarController.cs    # 農曆/節氣查詢
├── ChartController.cs       # 命盤圖表資料
└── CustomersController.cs   # 客戶資料管理

Data/
├── ApplicationDbContext.cs  # 主要 DbContext（Identity + 業務資料）
├── CalendarDbContext.cs     # 農曆資料專用 DbContext（同一 DB）
├── ApplicationUser.cs       # Identity 用戶（含生辰、點數欄位）
└── *.csv                    # 命理計算用靜態資料（編譯時複製至輸出）

Models/
├── DailyFortune.cs          # 每日運勢快取
├── FortuneRule.cs           # 運勢規則（知識庫）
├── KnowledgeDocument.cs     # 命理知識文件
├── Analysis/                # 紫微斗數查表資料 Model（StarStyle、PalaceMainStar 等）
└── ...

Services/
├── AstrologyService.cs      # 八字/紫微計算核心
├── AnalysisReportService.cs # 命書 DOCX 產生
├── ProReportService.cs      # 進階命書產生
├── BlindSchoolUltimateAnalyzer.cs # 盲派命理分析
├── YiZhuEngine.cs           # 易柱排盤引擎
├── ChartService.cs          # 命盤圖表
├── ExcelExportService.cs    # Excel 匯出
├── AstrologyEngine/         # 底層計算模組（常數、輔助函式）
└── Astrology/               # 神煞、地支關係計算

Migrations/                  # EF Core 遷移檔（ApplicationDbContext）
sql/                         # 生產環境手動執行的 SQL 腳本
wwwroot/
├── templates/               # DOCX 命書範本
└── images/                  # 靜態圖片
```

**注意：** `Services/BaziAnalysisService.cs` 與 `Services/IBaziAnalysisService.cs` 已在 `.csproj` 中排除編譯（`<Compile Remove=...>`），不屬於現行系統。

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
管理員判斷（`_config["Admin:Email"]` 比對）用於：
- `AstrologyController.analyze` — 專家命書下載
- `KnowledgeController` — 知識庫 CRUD
- `AdminController` — 用戶管理、點數調整
- `AuthController.profile` — `isAdmin` 欄位回傳

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

## 程式碼撰寫規範

- 程式碼只使用英文、數字、ASCII 符號
- 中文字只使用繁體中文
- 禁止在程式碼中使用 emoji 或 Unicode 特殊符號

## 注意事項

- Gemini API 呼叫 timeout 設為 5 分鐘（ConsultationController）、3 分鐘（FortuneController）
- 使用 `ILogger` 而非 `Console.WriteLine`（Fly.io logs 只顯示 ILogger 輸出）
- CORS 允許 `http://localhost:3000` 和 `https://myweb.fly.dev`
- 兩個 DbContext（`ApplicationDbContext`、`CalendarDbContext`）共用同一個 PostgreSQL 連線字串
- `Data/*.csv` 靜態資料在建置時複製至輸出目錄，供 `EcanService`/`AstrologyEngine` 讀取
