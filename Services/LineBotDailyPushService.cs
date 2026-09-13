using Ecanapi.Controllers;
using Ecanapi.Data;
using Ecanapi.Models;
using Microsoft.EntityFrameworkCore;
using System.Net.Http.Headers;
using System.Text.Json;

namespace Ecanapi.Services
{
    /// <summary>每天早上 7:30（台灣時間）推播九星開運通知給所有訂閱用戶</summary>
    public class LineBotDailyPushService : BackgroundService
    {
        private readonly IServiceScopeFactory _scopeFactory;
        private readonly IConfiguration _config;
        private readonly ILogger<LineBotDailyPushService> _logger;
        private static readonly HttpClient _httpClient = new();

        public LineBotDailyPushService(
            IServiceScopeFactory scopeFactory,
            IConfiguration config,
            ILogger<LineBotDailyPushService> logger)
        {
            _scopeFactory = scopeFactory;
            _config = config;
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                // 計算下一個台灣時間 7:30（UTC+8 = UTC 23:30 前一天）
                var nowUtc = DateTime.UtcNow;
                // 台灣時間 = UTC+8
                var nowTw = nowUtc.AddHours(8);
                var targetTw = nowTw.Date.AddHours(7).AddMinutes(30);
                if (nowTw >= targetTw) targetTw = targetTw.AddDays(1);
                var delay = targetTw - nowTw;

                _logger.LogInformation("LineBotDailyPush 下次推播：台灣時間 {Time}，等待 {Hours:F1} 小時",
                    targetTw.ToString("MM/dd HH:mm"), delay.TotalHours);

                try { await Task.Delay(delay, stoppingToken); }
                catch (OperationCanceledException) { break; }

                if (!stoppingToken.IsCancellationRequested)
                    await SendDailyPushAsync();
            }
        }

        private async Task SendDailyPushAsync()
        {
            _logger.LogInformation("LineBotDailyPush 開始推播...");
            try
            {
                using var scope = _scopeFactory.CreateScope();
                var context = scope.ServiceProvider.GetRequiredService<ApplicationDbContext>();
                var nineStarController = scope.ServiceProvider.GetRequiredService<NineStarController>();
                var fortuneController = scope.ServiceProvider.GetRequiredService<Controllers.FortuneController>();

                string accessToken = _config["LineBot:ChannelAccessToken"] ?? "";
                var pushDate = DateOnly.FromDateTime(DateTime.UtcNow.AddHours(8)); // 台灣時間日期

                // Block 1：LineUsers 九星推播（已在 LINE Bot 設定本命星的用戶，訂閱會員也收）
                var nineStarUsers = await context.LineUsers
                    .Where(u => u.NotifyEnabled && u.NatalStar > 0)
                    .ToListAsync();

                _logger.LogInformation("LineBotDailyPush 九星用戶推播 {Count} 位", nineStarUsers.Count);

                foreach (var lu in nineStarUsers)
                {
                    string message = "";
                    string status = "success";
                    string? errorMsg = null;
                    try
                    {
                        message = await nineStarController.NsBuildDailyFortune(lu.NatalStar);
                        await PushMessageAsync(accessToken, lu.LineUserId, message);
                        await Task.Delay(100);
                    }
                    catch (Exception ex)
                    {
                        status = "failed";
                        errorMsg = ex.Message;
                        _logger.LogError(ex, "LineBotDailyPush 九星用戶推播 {LineUserId} 失敗", lu.LineUserId);
                    }
                    context.LinePushLogs.Add(new LinePushLog
                    {
                        LineUserId = lu.LineUserId,
                        PushType = "ninestar",
                        ContentCategory = "ninestar-only",
                        NatalStar = lu.NatalStar,
                        BirthYear = lu.BirthYear > 0 ? lu.BirthYear : null,
                        PushDate = pushDate,
                        Message = message,
                        SentAt = DateTime.UtcNow,
                        Status = status,
                        ErrorMessage = errorMsg
                    });
                }
                await context.SaveChangesAsync();

                // Block 2：訂閱會員個人化推播（有效訂閱 + 已綁 LINE + 有出生年月日）
                var subscriberUsers = await context.UserSubscriptions
                    .Where(s => s.Status == "active" && s.ExpiryDate > DateTime.UtcNow)
                    .Join(context.Users, s => s.UserId, u => u.Id, (s, u) => u)
                    .Where(u => u.LineUserId != null
                             && u.BirthYear != null && u.BirthMonth != null
                             && u.BirthDay != null)
                    .Distinct()
                    .ToListAsync();

                // 過濾掉已在 LineUsers 表主動關閉通知的會員
                var disabledLineIds = await context.LineUsers
                    .Where(lu => !lu.NotifyEnabled)
                    .Select(lu => lu.LineUserId)
                    .ToListAsync();

                subscriberUsers = subscriberUsers
                    .Where(u => !disabledLineIds.Contains(u.LineUserId))
                    .ToList();

                _logger.LogInformation("LineBotDailyPush 訂閱會員推播 {Count} 位", subscriberUsers.Count);

                foreach (var subUser in subscriberUsers)
                {
                    string message = "";
                    string status = "success";
                    string? errorMsg = null;
                    int natalStar = 0;
                    string? shenShaHit = null;
                    try
                    {
                        string gender = subUser.BirthGender == 2 ? "F" : "M";
                        natalStar = NineStarController.NsCalcNatalStarStatic(
                            subUser.BirthYear!.Value, subUser.BirthMonth!.Value, subUser.BirthDay!.Value, gender);

                        // 優先使用八字個人化運勢（含喜用神、十神、神煞）
                        string? baziMsg = await fortuneController.BuildDailyPersonalFortuneText(subUser.Id);
                        if (!string.IsNullOrEmpty(baziMsg))
                        {
                            message = baziMsg;
                        }
                        else
                        {
                            // 無命盤資料時回退至九星運勢
                            message = await nineStarController.NsBuildDailyFortune(natalStar);
                        }

                        // 計算神煞（供日誌記錄）
                        var birthDt = new DateTime(subUser.BirthYear!.Value, subUser.BirthMonth!.Value, subUser.BirthDay!.Value);
                        string riZhu = GetGanZhiStem(birthDt);
                        string todayDiZhi = GetGanZhiBranch(DateTime.UtcNow.AddHours(8).Date); // 台灣時間
                        shenShaHit = Controllers.FortuneController.CalcShenSha(riZhu, todayDiZhi);

                        await PushMessageAsync(accessToken, subUser.LineUserId!, message);
                        await Task.Delay(100);
                    }
                    catch (Exception ex)
                    {
                        status = "failed";
                        errorMsg = ex.Message;
                        _logger.LogError(ex, "LineBotDailyPush 訂閱會員推播 {UserId} 失敗", subUser.Id);
                    }
                    context.LinePushLogs.Add(new LinePushLog
                    {
                        LineUserId = subUser.LineUserId!,
                        UserId = subUser.Id,
                        UserEmail = subUser.Email,
                        PushType = "subscriber",
                        ContentCategory = "bazi-enhanced",
                        NatalStar = natalStar,
                        BirthYear = subUser.BirthYear,
                        ShenShaHit = shenShaHit,
                        PushDate = pushDate,
                        Message = message,
                        SentAt = DateTime.UtcNow,
                        Status = status,
                        ErrorMessage = errorMsg
                    });
                }
                await context.SaveChangesAsync();

                _logger.LogInformation("LineBotDailyPush 推播完成");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "LineBotDailyPush 整批推播失敗");
            }
        }

        // 天干地支計算（同 FortuneController，避免跨 scope 呼叫副作用）
        private static readonly string[] Gan10 = {"甲","乙","丙","丁","戊","己","庚","辛","壬","癸"};
        private static readonly string[] Zhi12 = {"子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥"};
        private static readonly DateTime GanZhiEpoch = new DateTime(1900, 1, 31); // 甲辰日（地支辰=4，+4 修正）

        private static string GetGanZhiStem(DateTime date)
        {
            int days = (int)(date.Date - GanZhiEpoch.Date).TotalDays;
            return Gan10[((days % 10) + 10) % 10];
        }

        private static string GetGanZhiBranch(DateTime date)
        {
            int days = (int)(date.Date - GanZhiEpoch.Date).TotalDays;
            return Zhi12[(((days + 4) % 12) + 12) % 12]; // +4 修正地支偏移
        }

        private async Task PushMessageAsync(string accessToken, string lineUserId, string text)
        {
            var payload = new
            {
                to = lineUserId,
                messages = new[] { new { type = "text", text } }
            };
            var req = new HttpRequestMessage(HttpMethod.Post, "https://api.line.me/v2/bot/message/push");
            req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            req.Content = JsonContent.Create(payload);
            var resp = await _httpClient.SendAsync(req);
            if (!resp.IsSuccessStatusCode)
            {
                string body = await resp.Content.ReadAsStringAsync();
                throw new Exception($"LINE Push API 錯誤 {resp.StatusCode}: {body}");
            }
        }
    }
}
