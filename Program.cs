// 既有的 using 指示詞
using Ecanapi.Services;
using Ecanapi.Services.AstrologyEngine;
using Ecanapi.Data;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore;
using Microsoft.IdentityModel.Tokens;
using Microsoft.OpenApi.Models;
using Print2Engine;
using System.Text;
using Ecanapi.Models.Analysis;
using Microsoft.EntityFrameworkCore.Diagnostics;

var builder = WebApplication.CreateBuilder(args);

// --- 1. 註冊服務 ---
var connectionString = builder.Configuration.GetConnectionString("DefaultConnection");

builder.Services.AddDbContext<ApplicationDbContext>(options =>
{
    options.UseNpgsql(connectionString);
    options.ConfigureWarnings(w => w.Ignore(RelationalEventId.PendingModelChangesWarning));
});
builder.Services.AddDbContext<CalendarDbContext>(options =>
{
    options.UseNpgsql(connectionString);
    options.ConfigureWarnings(w => w.Ignore(RelationalEventId.PendingModelChangesWarning));
});

// 【⭐ 修正 1：放寬 Identity 密碼策略 (解決註冊問題) ⭐】
builder.Services.AddIdentity<ApplicationUser, IdentityRole>(options =>
{
    // 將密碼要求放寬到與您前端要求一致，以確保註冊能成功通過 Identity 的檢查
    options.Password.RequiredLength = 6;            // 最低長度 6 位數
    options.Password.RequireDigit = false;          // 不強制要求數字
    options.Password.RequireLowercase = false;      // 不強制要求小寫
    options.Password.RequireUppercase = false;      // 不強制要求大寫
    options.Password.RequireNonAlphanumeric = false; // 不強制要求特殊字符
    options.User.RequireUniqueEmail = true;         // 確保 Email 唯一性
})
    .AddEntityFrameworkStores<ApplicationDbContext>()
    .AddDefaultTokenProviders();

// JWT 認證配置
builder.Services.AddAuthentication(options =>
{
    options.DefaultAuthenticateScheme = JwtBearerDefaults.AuthenticationScheme;
    options.DefaultChallengeScheme = JwtBearerDefaults.AuthenticationScheme;
    options.DefaultScheme = JwtBearerDefaults.AuthenticationScheme;
})
.AddJwtBearer(options =>
{
    options.SaveToken = true;
    options.RequireHttpsMetadata = false;

    // 依您成功的版本，我們優先使用 JWT:Secret
    var jwtSecret = builder.Configuration["JWT:Secret"];

    // 【💥 修正 2：添加 Null 檢查 (解決 Swagger 崩潰問題) 💥】
    // 移除危險的 "!" 符號，並用 if 檢查取代，以防配置遺失導致運行時崩潰。
    if (string.IsNullOrEmpty(jwtSecret))
    {
        // 如果為 null，拋出更明確的錯誤，方便您排查 Fly.io 的環境變數設定
        throw new InvalidOperationException("JWT Secret 配置遺失。請確認 appsettings.json 或 Fly.io 的環境變數 JWT:Secret 是否正確設定。");
    }

    options.TokenValidationParameters = new TokenValidationParameters
    {
        ValidateIssuer = true,
        ValidateAudience = true,
        ValidateLifetime = true,
        ValidateIssuerSigningKey = true,
        ValidIssuer = builder.Configuration["JWT:ValidIssuer"],
        ValidAudience = builder.Configuration["JWT:ValidAudience"],
        // 使用我們檢查過且確保非 null 的 jwtSecret 變數
        IssuerSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(jwtSecret))
    };
});

builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();

builder.Services.AddSwaggerGen(options =>
{
    options.AddSecurityDefinition("Bearer", new OpenApiSecurityScheme
    {
        Description = "JWT Authorization header using the Bearer scheme. Example: \"Authorization: Bearer {token}\"",
        Name = "Authorization",
        In = ParameterLocation.Header,
        Type = SecuritySchemeType.Http,
        Scheme = "bearer"
    });

    options.AddSecurityRequirement(new OpenApiSecurityRequirement
    {
        {
            new OpenApiSecurityScheme
            {
                Reference = new OpenApiReference
                {
                    Type = ReferenceType.SecurityScheme,
                    Id = "Bearer"
                }
            },
            new string[] { }
        }
    });
});

// --- 註冊您所有的自訂服務 (保持不變) ---
builder.Services.AddScoped<Print2Engine.IEcanCalendar, EcanCalendarAdapter>();
builder.Services.AddScoped<Print2Engine.Print2Engine>();
builder.Services.AddScoped<ICalendarService, CalendarService>();
builder.Services.AddScoped<IChartService, ChartService>();
builder.Services.AddScoped<IAstrologyService, AstrologyService>();
builder.Services.AddScoped<IExcelExportService, ExcelExportService>();
builder.Services.AddScoped<IAnalysisService, AnalysisService>();
builder.Services.AddScoped<IAnalysisReportService, AnalysisReportService>();


// --- CORS 配置 (保持不變) ---
var AllowFrontendOrigins = "AllowFrontendOrigins";

builder.Services.AddCors(options =>
{
    options.AddPolicy(name: AllowFrontendOrigins, policy =>
    {
        policy.WithOrigins("http://localhost:3000",                  // 1. Next.js 本地開發端口
                           "https://myweb.fly.dev")                 // 2. 【新增】：您已部署的前端網址
             .AllowAnyMethod()
             .AllowAnyHeader()
             .AllowCredentials();
    });
});


var app = builder.Build();

// --- 2. 設定 HTTP 請求管線 ---
app.UseCors(AllowFrontendOrigins);

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();
app.UseAuthentication();
app.UseAuthorization();

app.MapControllers();

app.Run();