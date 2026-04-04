using Ecanapi.Data;
using Ecanapi.Services;
using Microsoft.EntityFrameworkCore;
using System.Net.Http.Headers;
using System.Text.Json;

namespace Ecanapi.Services
{
    /// <summary>每天早上 7:35（台灣時間）自動發佈今日九星運勢到 Instagram</summary>
    public class InstagramDailyPostService : BackgroundService
    {
        private readonly IServiceScopeFactory _scopeFactory;
        private readonly IConfiguration _config;
        private readonly ILogger<InstagramDailyPostService> _logger;
        private static readonly HttpClient _http = new();

        public InstagramDailyPostService(
            IServiceScopeFactory scopeFactory,
            IConfiguration config,
            ILogger<InstagramDailyPostService> logger)
        {
            _scopeFactory = scopeFactory;
            _config = config;
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                var nowTw = DateTime.UtcNow.AddHours(8);
                var targetTw = nowTw.Date.AddHours(7).AddMinutes(35);
                if (nowTw >= targetTw) targetTw = targetTw.AddDays(1);
                var delay = targetTw - nowTw;

                _logger.LogInformation("InstagramDailyPost 下次發佈：台灣時間 {Time}，等待 {Hours:F1} 小時",
                    targetTw.ToString("MM/dd HH:mm"), delay.TotalHours);

                try { await Task.Delay(delay, stoppingToken); }
                catch (OperationCanceledException) { break; }

                if (!stoppingToken.IsCancellationRequested)
                    var (_, _) = await PostDailyFortuneAsync();
            }
        }

        internal async Task<(bool ok, string message)> PostDailyFortuneAsync(string? overrideImageUrl = null)
        {
            _logger.LogInformation("InstagramDailyPost 開始發佈...");
            try
            {
                string accessToken = _config["IG_ACCESS_TOKEN"] ?? "";
                string igUserId = _config["IG_USER_ID"] ?? "";

                if (string.IsNullOrEmpty(accessToken) || string.IsNullOrEmpty(igUserId))
                {
                    _logger.LogWarning("InstagramDailyPost: IG_ACCESS_TOKEN 或 IG_USER_ID 未設定，略過");
                    return (false, "IG_ACCESS_TOKEN 或 IG_USER_ID 未設定");
                }

                string imageUrl = overrideImageUrl ?? await BuildImageUrlAsync();
                _logger.LogInformation("InstagramDailyPost 圖片 URL: {Url}", imageUrl);

                string caption = await BuildCaptionAsync();

                // Step 1: 建立媒體容器
                var (creationId, createError) = await CreateMediaContainerAsync(igUserId, accessToken, imageUrl, caption);
                if (string.IsNullOrEmpty(creationId))
                {
                    _logger.LogError("InstagramDailyPost: 建立媒體容器失敗：{Error}", createError);
                    return (false, $"建立媒體容器失敗：{createError}");
                }

                // 等待容器處理
                await Task.Delay(3000);

                // Step 2: 發佈
                await PublishMediaAsync(igUserId, accessToken, creationId);
                _logger.LogInformation("InstagramDailyPost 發佈成功，圖片：{Url}", imageUrl);
                return (true, $"發佈成功！圖片主題：{imageUrl.Split('?')[0]}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "InstagramDailyPost 發佈失敗");
                return (false, ex.Message);
            }
        }

        private Task<string> BuildImageUrlAsync()
        {
            // 使用 /api/ig-card/{date}.png 格式（Instagram 要求 URL 需以 .png/.jpg 結尾）
            var nowTw = DateTime.UtcNow.AddHours(8);
            string date = nowTw.ToString("yyyy-MM-dd");
            return Task.FromResult($"https://myweb.fly.dev/api/ig-card/{date}.jpg");
        }

        private async Task<string> BuildCaptionAsync()
        {
            var nowTw = DateTime.UtcNow.AddHours(8);
            int dayStar = NineStarCalcHelper.CalcDayStar(nowTw);
            int yearStar = NineStarCalcHelper.CalcYearStar(nowTw.Year);
            int monthStar = NineStarCalcHelper.CalcMonthStar(nowTw);

            string dayStarName = NineStarCalcHelper.StarNames[dayStar];
            string dayStarDir = NineStarCalcHelper.StarDirections[dayStar];
            string dayStarColor = NineStarCalcHelper.StarColors[dayStar];

            // 嘗試從 DB 取得今日運勢文字（用流日星自我組合）
            string fortuneText = "";
            string auspicious = "";
            string avoid = "";
            try
            {
                using var scope = _scopeFactory.CreateScope();
                var context = scope.ServiceProvider.GetRequiredService<ApplicationDbContext>();
                var rule = await context.NineStarDailyRules
                    .FirstOrDefaultAsync(r => r.NatalStar == dayStar && r.FlowStar == dayStar);
                if (rule != null && !string.IsNullOrEmpty(rule.FortuneText))
                {
                    fortuneText = rule.FortuneText;
                    auspicious = rule.Auspicious ?? "";
                    avoid = rule.Avoid ?? "";
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "InstagramDailyPost: 取得運勢文字失敗，使用預設文字");
            }

            if (string.IsNullOrEmpty(fortuneText))
                fortuneText = $"今日流日{dayStarName}，把握吉方位{dayStarDir}，以幸運色{dayStarColor}提升能量。";

            string caption = $"【{nowTw:MM/dd} 今日九星開運】\n\n" +
                $"流年：{NineStarCalcHelper.StarNames[yearStar]}　" +
                $"流月：{NineStarCalcHelper.StarNames[monthStar]}　" +
                $"流日：{dayStarName}\n\n" +
                $"{fortuneText}\n\n";

            if (!string.IsNullOrEmpty(auspicious)) caption += $"宜：{auspicious}\n";
            if (!string.IsNullOrEmpty(avoid)) caption += $"忌：{avoid}\n";

            caption += $"吉方位：{dayStarDir}　幸運色：{dayStarColor}\n\n" +
                $"個人化八字命盤分析 ➡ myweb.fly.dev\n\n" +
                $"#玉洞子星相古學堂 #命理 #八字 #九星氣學 #每日運勢 #開運 #命盤 #紫微斗數 #風水";

            return caption;
        }

        private async Task<(string creationId, string error)> CreateMediaContainerAsync(string userId, string token, string imageUrl, string caption)
        {
            var url = $"https://graph.facebook.com/v21.0/{userId}/media";
            var payload = new Dictionary<string, string>
            {
                ["image_url"] = imageUrl,
                ["media_type"] = "IMAGE",
                ["caption"] = caption,
                ["access_token"] = token
            };

            var req = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new FormUrlEncodedContent(payload)
            };

            var resp = await _http.SendAsync(req);
            string body = await resp.Content.ReadAsStringAsync();

            if (!resp.IsSuccessStatusCode)
            {
                _logger.LogError("InstagramDailyPost CreateContainer 失敗 {Status}: {Body}", resp.StatusCode, body);
                return ("", $"HTTP {(int)resp.StatusCode}: {body[..Math.Min(body.Length, 300)]}");
            }

            using var doc = JsonDocument.Parse(body);
            string id = doc.RootElement.TryGetProperty("id", out var idProp) ? idProp.GetString() ?? "" : "";
            return (id, string.IsNullOrEmpty(id) ? "回傳無 id 欄位" : "");
        }

        private async Task PublishMediaAsync(string userId, string token, string creationId)
        {
            var url = $"https://graph.facebook.com/v21.0/{userId}/media_publish";
            var payload = new Dictionary<string, string>
            {
                ["creation_id"] = creationId,
                ["access_token"] = token
            };

            var req = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new FormUrlEncodedContent(payload)
            };

            var resp = await _http.SendAsync(req);
            if (!resp.IsSuccessStatusCode)
            {
                string body = await resp.Content.ReadAsStringAsync();
                throw new Exception($"Instagram Publish 失敗 {resp.StatusCode}: {body}");
            }
        }
    }
}
