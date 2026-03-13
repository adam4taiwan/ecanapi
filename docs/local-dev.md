# 本機開發啟動指南（Swagger on localhost:5013）

## 前置條件

- .NET 8 SDK 已安裝
- PostgreSQL 正在執行（Windows 本機，Port 5432）
- `appsettings.Development.json` 存在（已加入 .gitignore，需手動建立）

---

## 步驟一：確認 `appsettings.Development.json`

在專案根目錄建立 `appsettings.Development.json`，WSL2 環境下 PostgreSQL host 為 Windows 端 IP（通常是 `172.22.192.1` 或 `$(cat /etc/resolv.conf | grep nameserver | awk '{print $2}')`）：

```json
{
  "ConnectionStrings": {
    "DefaultConnection": "Host=172.22.192.1;Port=5432;Database=EcanApiDb;Username=postgres;Password=<your-password>"
  },
  "Logging": {
    "LogLevel": {
      "Default": "Information",
      "Microsoft.AspNetCore": "Warning"
    }
  },
  "AllowedHosts": "*",
  "JWT": {
    "ValidAudience": "your-flutter-app",
    "ValidIssuer": "your-backend-api",
    "Secret": "<至少32字元的隨機字串>"
  },
  "Gemini": {
    "ApiKey": "<your-gemini-api-key>"
  },
  "Stripe": {
    "SecretKey": "<your-stripe-secret-key>",
    "WebhookSecret": "<your-stripe-webhook-secret>"
  },
  "Admin": {
    "Email": "adam4taiwan@gmail.com"
  }
}
```

> 若 Windows 端 PostgreSQL 需開放 WSL2 連線，確認 `pg_hba.conf` 允許來自 `172.22.0.0/16` 的連線，並在 `postgresql.conf` 設定 `listen_addresses = '*'`。

---

## 步驟二：確認 PostgreSQL 資料庫存在

在 PostgreSQL 建立資料庫（若尚未存在）：

```sql
CREATE DATABASE "EcanApiDb";
```

---

## 步驟三：執行 Migration（首次或有新 Migration 時）

```bash
export PATH="$PATH:$HOME/.dotnet/tools"
cd /home/adamtsai/projects/Ecanapi
dotnet ef database update --context ApplicationDbContext
```

> 本機 `IsDevelopment()` 環境下，`dotnet run` 啟動時也會自動執行 `Migrate()`，此步驟可省略。

---

## 步驟四：啟動 API

```bash
cd /home/adamtsai/projects/Ecanapi
dotnet run --launch-profile http
```

啟動後輸出會顯示：

```
Now listening on: http://0.0.0.0:5013
```

---

## 步驟五：開啟 Swagger UI

瀏覽器前往：

```
http://localhost:5013/swagger
```

---

## 測試 JWT 認證

1. 使用 `POST /api/Auth/register` 註冊帳號
2. 使用 `POST /api/Auth/login` 取得 Token
3. 點擊 Swagger 右上角 **Authorize** 按鈕
4. 輸入：`Bearer <取得的Token>`
5. 確認後即可呼叫需要認證的端點

---

## 常見問題

| 問題 | 解法 |
|------|------|
| `JWT Secret 配置遺失` | 確認 `appsettings.Development.json` 有 `JWT:Secret` 且長度足夠 |
| 連線 PostgreSQL 失敗 | 確認 Windows PostgreSQL 已啟動，且 `172.22.192.1` 為當前 WSL2 的 host IP |
| Swagger 頁面 404 | 確認 `ASPNETCORE_ENVIRONMENT=Development`（`--launch-profile http` 已設定） |
| Migration 失敗 | 確認資料庫存在，且連線字串正確 |
