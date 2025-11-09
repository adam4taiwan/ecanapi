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

var builder = WebApplication.CreateBuilder(args);

// --- 1. 註冊服務 ---
var connectionString = builder.Configuration.GetConnectionString("DefaultConnection");

builder.Services.AddDbContext<ApplicationDbContext>(options =>
    options.UseNpgsql(connectionString));
builder.Services.AddDbContext<CalendarDbContext>(options =>
    options.UseNpgsql(connectionString));

builder.Services.AddIdentity<ApplicationUser, IdentityRole>()
    .AddEntityFrameworkStores<ApplicationDbContext>()
    .AddDefaultTokenProviders();

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
    options.TokenValidationParameters = new TokenValidationParameters
    {
        ValidateIssuer = true,
        ValidateAudience = true,
        ValidateLifetime = true,
        ValidateIssuerSigningKey = true,
        ValidIssuer = builder.Configuration["JWT:ValidIssuer"],
        ValidAudience = builder.Configuration["JWT:ValidAudience"],
        IssuerSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(builder.Configuration["JWT:Secret"]!))
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

// --- 註冊您所有的自訂服務 ---
builder.Services.AddScoped<Print2Engine.IEcanCalendar, EcanCalendarAdapter>();
builder.Services.AddScoped<Print2Engine.Print2Engine>();
builder.Services.AddScoped<ICalendarService, CalendarService>();
builder.Services.AddScoped<IChartService, ChartService>();
builder.Services.AddScoped<IAstrologyService, AstrologyService>();
builder.Services.AddScoped<IExcelExportService, ExcelExportService>();
// 【新增】註冊新的分析服務
builder.Services.AddScoped<IAnalysisService, AnalysisService>();
// 【新增】註冊新的斷命分析報告服務
builder.Services.AddScoped<IAnalysisReportService, AnalysisReportService>();

// --- 更新 CORS 配置 ---
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowNextJsLocal", policy =>
    {
        policy.WithOrigins("http://localhost:3000")  // 允許 Next.js 本地開發端口
              .AllowAnyMethod()  // 允許 GET, POST 等所有方法
              .AllowAnyHeader()  // 允許所有標頭（如 Content-Type, Authorization）
              .AllowCredentials();  // 支援 JWT 認證（如需要 Cookie 或認證）
    });
});

var app = builder.Build();

// --- 2. 設定 HTTP 請求管線 ---
app.UseCors("AllowNextJsLocal");  // 應用新的 CORS 政策

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