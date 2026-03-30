using Ecanapi.Controllers;
using Ecanapi.Data;
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
                var controller = scope.ServiceProvider.GetRequiredService<NineStarController>();

                // 取得所有開啟通知且已設定本命星的用戶
                var users = await context.LineUsers
                    .Where(u => u.NotifyEnabled && u.NatalStar > 0)
                    .ToListAsync();

                _logger.LogInformation("LineBotDailyPush 推播 {Count} 位用戶", users.Count);

                string accessToken = _config["LineBot:ChannelAccessToken"] ?? "";

                foreach (var user in users)
                {
                    try
                    {
                        string message = await controller.NsBuildDailyFortune(user.NatalStar);
                        await PushMessageAsync(accessToken, user.LineUserId, message);
                        await Task.Delay(100); // 避免 LINE API rate limit
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "LineBotDailyPush 推播用戶 {UserId} 失敗", user.LineUserId);
                    }
                }

                _logger.LogInformation("LineBotDailyPush 推播完成");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "LineBotDailyPush 整批推播失敗");
            }
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
