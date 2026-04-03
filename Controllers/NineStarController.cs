using Ecanapi.Data;
using Ecanapi.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Net.Http.Headers;
using System.Security.Claims;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace Ecanapi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class NineStarController : ControllerBase
    {
        private readonly ApplicationDbContext _context;
        private readonly IConfiguration _config;
        private readonly ILogger<NineStarController> _logger;
        private static readonly HttpClient _httpClient = new HttpClient { Timeout = TimeSpan.FromMinutes(3) };

        // ============================================================
        //  靜態資料表
        // ============================================================
        private static readonly string[] StarNames =
        {
            "", "一白水星", "二黑土星", "三碧木星", "四綠木星", "五黃土星",
            "六白金星", "七赤金星", "八白土星", "九紫火星"
        };

        // 吉方位（1-9）
        private static readonly string[] StarDirections =
        {
            "", "北方", "西南方", "東方", "東南方", "中宮（避中央）", "西北方", "西方", "東北方", "南方"
        };

        // 吉顏色（1-9）
        private static readonly string[] StarColors =
        {
            "", "白色、藍色", "黃色、棕色", "綠色、青色", "綠色、青色", "黃色（需化解）",
            "白色、金色", "白色、金色", "白色、黃色", "紫色、紅色"
        };

        // 幸運數字（1-9）
        private static readonly int[] StarNumbers = { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9 };

        // 甲子參考日：January 1, 2000 = 甲戌（60循環 index 10）
        // 甲子（index 0）= Jan 1, 2000 - 10 days = Dec 22, 1999
        private static readonly DateTime NsEpochDate = new DateTime(2000, 1, 1);
        private const int NsEpochCycleIndex = 10; // Jan 1, 2000 = 甲戌 (60-cycle index 10)

        // ============================================================
        //  建構子
        // ============================================================
        public NineStarController(ApplicationDbContext context, IConfiguration config, ILogger<NineStarController> logger)
        {
            _context = context;
            _config = config;
            _logger = logger;
        }

        // ============================================================
        //  公開端點
        // ============================================================

        /// <summary>計算本命星（公開，無需登入）</summary>
        [HttpGet("natal")]
        public IActionResult GetNatalStar(int year, int month, int day, string gender)
        {
            if (year < 1900 || year > 2100) return BadRequest("年份超出範圍");
            if (gender != "M" && gender != "F") return BadRequest("gender 須為 M 或 F");

            int star = NsCalcNatalStar(year, month, day, gender);
            return Ok(new
            {
                natalStar = star,
                starName = StarNames[star],
                direction = StarDirections[star],
                color = StarColors[star],
                luckyNumber = StarNumbers[star]
            });
        }

        /// <summary>取得今日年/月/日/時四星（公開）</summary>
        [HttpGet("stars/today")]
        public IActionResult GetTodayStars()
        {
            var now = DateTime.Now;
            int yearStar = NsCalcYearStar(now.Year);
            int monthStar = NsCalcMonthStar(now);
            int dayStar = NsCalcDayStar(now);
            int hourStar = NsCalcHourStar(now);

            return Ok(new
            {
                date = now.ToString("yyyy-MM-dd HH:mm"),
                yearStar = new { number = yearStar, name = StarNames[yearStar] },
                monthStar = new { number = monthStar, name = StarNames[monthStar] },
                dayStar = new { number = dayStar, name = StarNames[dayStar] },
                hourStar = new { number = hourStar, name = StarNames[hourStar] }
            });
        }

        /// <summary>取得某顆本命星的特質（KB 優先，空則 Gemini 生成後回填）</summary>
        [HttpGet("trait/{starNumber}")]
        public async Task<IActionResult> GetTrait(int starNumber)
        {
            if (starNumber < 1 || starNumber > 9) return BadRequest("星號須為 1-9");

            var trait = await _context.NineStarTraits.FirstOrDefaultAsync(t => t.StarNumber == starNumber);

            // 若不存在，先建立空記錄
            if (trait == null)
            {
                trait = new NineStarTrait
                {
                    StarNumber = starNumber,
                    StarName = StarNames[starNumber],
                    LuckyDirection = StarDirections[starNumber],
                    LuckyColor = StarColors[starNumber],
                    LuckyNumber = StarNumbers[starNumber]
                };
                _context.NineStarTraits.Add(trait);
                await _context.SaveChangesAsync();
            }

            // KB 缺 Personality → Gemini 生成後回填
            if (string.IsNullOrEmpty(trait.Personality))
            {
                trait.Personality = await NsGeminiGenTrait(starNumber, "personality");
                trait.Career = await NsGeminiGenTrait(starNumber, "career");
                trait.Relationship = await NsGeminiGenTrait(starNumber, "relationship");
                trait.Health = await NsGeminiGenTrait(starNumber, "health");
                trait.UpdatedAt = DateTime.UtcNow;
                await _context.SaveChangesAsync();
                _logger.LogInformation("NineStar trait {Star} generated by Gemini and saved to KB", starNumber);
            }

            return Ok(new
            {
                starNumber = trait.StarNumber,
                starName = trait.StarName,
                personality = trait.Personality,
                career = trait.Career,
                relationship = trait.Relationship,
                health = trait.Health,
                luckyDirection = trait.LuckyDirection,
                luckyColor = trait.LuckyColor,
                luckyNumber = trait.LuckyNumber
            });
        }

        /// <summary>LINE Bot 用戶登記生辰（無需認證）</summary>
        /// <summary>會員中心設定每日 LINE 推播通知（需登入，依 AspNetUsers.LineUserId 找到 LineUser）</summary>
        [HttpPut("notify")]
        [Authorize]
        public async Task<IActionResult> SetNotify([FromBody] SetNotifyRequest req)
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            var appUser = await _context.Users.FindAsync(userId);
            if (appUser?.LineUserId == null)
                return BadRequest(new { error = "請先用 LINE 帳號登入本平台，才能設定 LINE 推播通知" });

            var lineUser = await _context.LineUsers.FirstOrDefaultAsync(u => u.LineUserId == appUser.LineUserId);
            if (lineUser == null)
                return BadRequest(new { error = "尚未建立 LINE Bot 用戶資料，請先加入官方帳號" });

            lineUser.NotifyEnabled = req.Enabled;
            lineUser.UpdatedAt = DateTime.UtcNow;
            await _context.SaveChangesAsync();

            return Ok(new { notifyEnabled = lineUser.NotifyEnabled, message = req.Enabled ? "已開啟每日 LINE 推播" : "已關閉每日 LINE 推播" });
        }

        /// <summary>查詢目前推播設定狀態（需登入）</summary>
        [HttpGet("notify")]
        [Authorize]
        public async Task<IActionResult> GetNotify()
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            var appUser = await _context.Users.FindAsync(userId);
            if (appUser?.LineUserId == null)
                return Ok(new { hasLineLinked = false, notifyEnabled = false });

            var lineUser = await _context.LineUsers.FirstOrDefaultAsync(u => u.LineUserId == appUser.LineUserId);
            return Ok(new
            {
                hasLineLinked = true,
                hasBotRecord = lineUser != null,
                notifyEnabled = lineUser?.NotifyEnabled ?? false,
                natalStar = lineUser?.NatalStar ?? 0
            });
        }

        [HttpPost("register")]
        public async Task<IActionResult> Register([FromBody] LineRegisterRequest req)
        {
            if (string.IsNullOrEmpty(req.LineUserId)) return BadRequest("LineUserId 必填");
            if (req.BirthYear < 1900 || req.BirthYear > 2100) return BadRequest("年份超出範圍");
            if (req.Gender != "M" && req.Gender != "F") return BadRequest("gender 須為 M 或 F");

            int natalStar = NsCalcNatalStar(req.BirthYear, req.BirthMonth, req.BirthDay, req.Gender);

            var existing = await _context.LineUsers.FirstOrDefaultAsync(u => u.LineUserId == req.LineUserId);
            if (existing != null)
            {
                existing.BirthYear = req.BirthYear;
                existing.BirthMonth = req.BirthMonth;
                existing.BirthDay = req.BirthDay;
                existing.Gender = req.Gender;
                existing.NatalStar = natalStar;
                existing.DisplayName = req.DisplayName;
                existing.UpdatedAt = DateTime.UtcNow;
            }
            else
            {
                _context.LineUsers.Add(new LineUser
                {
                    LineUserId = req.LineUserId,
                    BirthYear = req.BirthYear,
                    BirthMonth = req.BirthMonth,
                    BirthDay = req.BirthDay,
                    Gender = req.Gender,
                    NatalStar = natalStar,
                    DisplayName = req.DisplayName
                });
            }
            await _context.SaveChangesAsync();

            return Ok(new
            {
                natalStar,
                starName = StarNames[natalStar],
                message = "登記成功"
            });
        }

        // ============================================================
        //  Admin 手動測試推播
        // ============================================================

        /// <summary>Admin 手動觸發每日推播（立即執行，用於測試驗證）</summary>
        [HttpPost("push-now")]
        [Authorize]
        public async Task<IActionResult> PushNow()
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            var user = await _context.Users.FindAsync(userId);
            var adminEmail = _config["Admin:Email"];
            if (user == null || !string.Equals(user.Email, adminEmail, StringComparison.OrdinalIgnoreCase))
                return Forbid();

            var accessToken = _config["LineBot:ChannelAccessToken"] ?? "";
            var fortuneController = HttpContext.RequestServices.GetRequiredService<FortuneController>();
            var pushed = new List<object>();
            var errors = new List<object>();

            // Block 1：九星推播
            var nineStarUsers = await _context.LineUsers
                .Where(u => u.NotifyEnabled && u.NatalStar > 0)
                .ToListAsync();

            foreach (var lu in nineStarUsers)
            {
                try
                {
                    string message = await NsBuildDailyFortune(lu.NatalStar);
                    await NsPushMessage(accessToken, lu.LineUserId, message);
                    pushed.Add(new { type = "nineStar", lineUserId = lu.LineUserId, name = lu.DisplayName });
                    await Task.Delay(100);
                }
                catch (Exception ex)
                {
                    errors.Add(new { type = "nineStar", lineUserId = lu.LineUserId, error = ex.Message });
                }
            }

            // Block 2：訂閱會員個人化推播
            var subscriberUsers = await _context.UserSubscriptions
                .Where(s => s.Status == "active" && s.ExpiryDate > DateTime.UtcNow)
                .Join(_context.Users, s => s.UserId, u => u.Id, (s, u) => u)
                .Where(u => u.LineUserId != null
                         && u.BirthYear != null && u.BirthMonth != null
                         && u.BirthDay != null && u.BirthHour != null)
                .Distinct()
                .ToListAsync();

            foreach (var subUser in subscriberUsers)
            {
                try
                {
                    string? message = await fortuneController.BuildDailyPersonalFortuneText(subUser.Id);
                    if (string.IsNullOrEmpty(message)) continue;
                    await NsPushMessage(accessToken, subUser.LineUserId!, message);
                    pushed.Add(new { type = "subscriber", lineUserId = subUser.LineUserId, name = subUser.Name });
                    await Task.Delay(100);
                }
                catch (Exception ex)
                {
                    errors.Add(new { type = "subscriber", userId = subUser.Id, error = ex.Message });
                }
            }

            return Ok(new
            {
                pushedCount = pushed.Count,
                errorCount = errors.Count,
                pushed,
                errors
            });
        }

        private async Task NsPushMessage(string accessToken, string lineUserId, string text)
        {
            var payload = new { to = lineUserId, messages = new[] { new { type = "text", text } } };
            var req = new HttpRequestMessage(HttpMethod.Post, "https://api.line.me/v2/bot/message/push");
            req.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
            req.Content = JsonContent.Create(payload);
            var resp = await _httpClient.SendAsync(req);
            if (!resp.IsSuccessStatusCode)
            {
                string body = await resp.Content.ReadAsStringAsync();
                throw new Exception($"LINE Push API 錯誤 {resp.StatusCode}: {body}");
            }
        }

        // ============================================================
        //  LINE Bot Webhook
        // ============================================================

        /// <summary>LINE Bot Webhook（接收 LINE 訊息事件）</summary>
        [HttpPost("webhook")]
        public async Task<IActionResult> Webhook()
        {
            // 讀取原始 body（驗簽用）
            using var reader = new StreamReader(Request.Body, Encoding.UTF8);
            string body = await reader.ReadToEndAsync();

            // 驗證 X-Line-Signature
            string sig = Request.Headers["X-Line-Signature"].FirstOrDefault() ?? "";
            if (!NsValidateLineSig(body, sig))
            {
                _logger.LogWarning("LINE Webhook 簽名驗證失敗");
                return Unauthorized();
            }

            var json = JsonSerializer.Deserialize<JsonElement>(body);
            if (!json.TryGetProperty("events", out var events)) return Ok();

            foreach (var evt in events.EnumerateArray())
            {
                string evtType = evt.GetProperty("type").GetString() ?? "";
                string replyToken = evt.TryGetProperty("replyToken", out var rt) ? rt.GetString() ?? "" : "";
                string lineUserId = evt.TryGetProperty("source", out var src)
                    ? src.TryGetProperty("userId", out var uid) ? uid.GetString() ?? "" : ""
                    : "";

                if (evtType == "follow")
                {
                    bool subscribed = await NsIsSubscribed(lineUserId);
                    await NsLineReply(replyToken, subscribed ? NsWelcomeText() : NsNotSubscribedText());
                }
                else if (evtType == "message")
                {
                    string msgType = evt.GetProperty("message").GetProperty("type").GetString() ?? "";
                    if (msgType == "text")
                    {
                        bool subscribed = await NsIsSubscribed(lineUserId);
                        if (!subscribed)
                        {
                            await NsLineReply(replyToken, NsNotSubscribedText());
                            continue;
                        }
                        string text = evt.GetProperty("message").GetProperty("text").GetString()?.Trim() ?? "";
                        string reply = await NsHandleText(text, lineUserId);
                        if (!string.IsNullOrEmpty(reply))
                            await NsLineReply(replyToken, reply);
                    }
                }
            }

            return Ok();
        }

        // ── 狀態機主入口 ─────────────────────────────────────────────────

        private async Task<string> NsHandleText(string text, string lineUserId)
        {
            // 0 = 隨時回選單
            if (text == "0")
                return NsMenuText();

            var lineUser = await _context.LineUsers.FirstOrDefaultAsync(u => u.LineUserId == lineUserId);
            return await NsHandleIdle(text, lineUserId, lineUser);
        }

        // ── Idle 狀態：處理數字選單 ──────────────────────────────────────

        private async Task<string> NsHandleIdle(string text, string lineUserId, LineUser? lineUser)
        {
            switch (text)
            {
                case "1": // 每日通知 開/關
                    return await NsToggleNotify(lineUserId, lineUser);

                default:
                    return NsMenuText();
            }
        }

        // ── 以下功能暫時停用（保留供未來恢復）──────────────────────────
        /*
        private async Task<string> NsAutoStartReg(string lineUserId, string featureName) { ... }
        private async Task<string> NsRegYear(...) { ... }
        private async Task<string> NsRegMonth(...) { ... }
        private async Task<string> NsRegDay(...) { ... }
        private async Task<string> NsRegGender(...) { ... }
        private async Task<string> NsBuildPersonality(int star) { ... }
        */

        // ── 每日通知訂閱切換 ─────────────────────────────────────────────

        private async Task<string> NsToggleNotify(string lineUserId, LineUser? lineUser)
        {
            if (lineUser == null)
            {
                // 首次操作：建立記錄並開啟通知
                _context.LineUsers.Add(new LineUser
                {
                    LineUserId = lineUserId, State = "idle", NotifyEnabled = true
                });
                await _context.SaveChangesAsync();
                return "已開啟每日開運通知！\n每天早上 7:30 自動推送今日個人化開運建議。\n\n輸入 1 可隨時關閉。";
            }

            lineUser.NotifyEnabled = !lineUser.NotifyEnabled;
            lineUser.UpdatedAt = DateTime.UtcNow;
            await _context.SaveChangesAsync();

            return lineUser.NotifyEnabled
                ? "已開啟每日開運通知！\n每天早上 7:30 自動推送今日個人化開運建議。\n\n輸入 1 可隨時關閉。"
                : "已關閉每日開運通知。\n\n輸入 1 可重新開啟。";
        }

        // ── 共用：建立今日運勢訊息 ──────────────────────────────────────

        internal async Task<string> NsBuildDailyFortune(int natalStar)
        {
            var now = DateTime.UtcNow.AddHours(8); // Taiwan time
            int dayStar   = NsCalcDayStar(now);
            int yearStar  = NsCalcYearStar(now.Year);
            int monthStar = NsCalcMonthStar(now);

            var rule = await _context.NineStarDailyRules
                .FirstOrDefaultAsync(r => r.NatalStar == natalStar && r.FlowStar == dayStar);
            if (rule == null || string.IsNullOrEmpty(rule.FortuneText))
            {
                var gen = await NsGeminiGenDailyAdvice(natalStar, dayStar);
                if (rule == null) { rule = new NineStarDailyRule { NatalStar = natalStar, FlowStar = dayStar }; _context.NineStarDailyRules.Add(rule); }
                rule.FortuneText = gen.FortuneText; rule.Auspicious = gen.Auspicious;
                rule.Avoid = gen.Avoid; rule.Direction = gen.Direction; rule.Color = gen.Color;
                rule.UpdatedAt = DateTime.UtcNow;
                await _context.SaveChangesAsync();
            }
            return $"【{now:MM/dd} 九星開運】\n本命：{StarNames[natalStar]}\n流年：{StarNames[yearStar]} 流月：{StarNames[monthStar]} 流日：{StarNames[dayStar]}\n\n{rule.FortuneText}\n\n宜：{rule.Auspicious}\n忌：{rule.Avoid}\n吉方位：{rule.Direction ?? StarDirections[natalStar]}\n幸運色：{rule.Color ?? StarColors[natalStar]}";
        }

        private async Task<string> NsBuildPersonality(int star)
        {
            var trait = await _context.NineStarTraits.FirstOrDefaultAsync(t => t.StarNumber == star);
            if (trait == null || string.IsNullOrEmpty(trait.Personality))
            {
                string p = await NsGeminiGenTrait(star, "personality");
                if (trait == null)
                {
                    trait = new NineStarTrait { StarNumber = star, StarName = StarNames[star],
                        LuckyDirection = StarDirections[star], LuckyColor = StarColors[star], LuckyNumber = StarNumbers[star] };
                    _context.NineStarTraits.Add(trait);
                }
                trait.Personality = p; trait.UpdatedAt = DateTime.UtcNow;
                await _context.SaveChangesAsync();
            }
            string text = trait.Personality ?? "";
            if (text.Length > 300) text = text[..300] + "...";
            return $"【{StarNames[star]} 個性特質】\n\n{text}\n\n吉方位：{StarDirections[star]}\n幸運色：{StarColors[star]}";
        }

        // ── 狀態機輔助 ──────────────────────────────────────────────────

        private async Task NsSetState(string lineUserId, string state)
        {
            var user = await _context.LineUsers.FirstOrDefaultAsync(u => u.LineUserId == lineUserId);
            if (user != null) { user.State = state; user.UpdatedAt = DateTime.UtcNow; await _context.SaveChangesAsync(); }
        }

        private async Task NsSaveTempAndState(string lineUserId, LineUser? lineUser, string newState,
            int? year, int? month, int? day)
        {
            if (lineUser == null) return;
            if (year != null)  lineUser.TempYear  = year;
            if (month != null) lineUser.TempMonth = month;
            if (day != null)   lineUser.TempDay   = day;
            lineUser.State = newState;
            lineUser.UpdatedAt = DateTime.UtcNow;
            await _context.SaveChangesAsync();
        }

        private bool NsValidateLineSig(string body, string signature)
        {
            var secret = _config["LineBot:ChannelSecret"] ?? "";
            using var hmac = new HMACSHA256(Encoding.UTF8.GetBytes(secret));
            var hash = hmac.ComputeHash(Encoding.UTF8.GetBytes(body));
            return Convert.ToBase64String(hash) == signature;
        }

        private async Task NsLineReply(string replyToken, string text)
        {
            if (string.IsNullOrEmpty(replyToken)) return;
            var token = _config["LineBot:ChannelAccessToken"] ?? "";
            var payload = new { replyToken, messages = new[] { new { type = "text", text } } };
            var req = new HttpRequestMessage(HttpMethod.Post, "https://api.line.me/v2/bot/message/reply");
            req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
            req.Content = JsonContent.Create(payload);
            await _httpClient.SendAsync(req);
        }

        // ── 訂閱驗證 ─────────────────────────────────────────────────────

        /// <summary>檢查 LINE userId 是否綁定有效的 MyWeb 訂閱會員</summary>
        private async Task<bool> NsIsSubscribed(string lineUserId)
        {
            if (string.IsNullOrEmpty(lineUserId)) return false;
            var appUser = await _context.Users.FirstOrDefaultAsync(u => u.LineUserId == lineUserId);
            if (appUser == null) return false;
            var now = DateTime.UtcNow;
            return await _context.UserSubscriptions.AnyAsync(s =>
                s.UserId == appUser.Id &&
                s.Status == "active" &&
                s.ExpiryDate > now);
        }

        private static string NsNotSubscribedText() =>
            "您好！玉洞子每日開運建議為訂閱會員專屬服務。\n\n請先前往以下網址加入訂閱會員：\nhttps://yudongzi.tw/\n\n訂閱後即可享有每日個人化開運建議！";

        private static string NsWelcomeText() =>
            "歡迎加入【玉洞子星相古學堂】！\n\n您是訂閱會員，每天早上 7:30 將自動推播今日個人化開運建議。\n\n" + NsMenuText();

        private static string NsMenuText() =>
            "━━ 功能選單 ━━\n1. 每日開運通知（開/關）\n0. 顯示此選單";

        /// <summary>個人化今日九星建議（LINE Bot 用 lineUserId；網頁用 JWT）</summary>
        [HttpGet("daily")]
        public async Task<IActionResult> GetDaily(string? lineUserId = null)
        {
            int natalStar = 0;
            string birthStem = ""; // 出生年干（供十神推算）

            // 優先用 JWT 取本命星
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (!string.IsNullOrEmpty(userId))
            {
                var user = await _context.Users.FindAsync(userId);
                if (user != null && user.HasBirthData)
                {
                    natalStar = NsCalcNatalStar(
                        user.BirthYear!.Value, user.BirthMonth!.Value, user.BirthDay!.Value,
                        user.BirthGender == 2 ? "F" : "M");
                    birthStem = Ecanapi.Services.NineStarCalcHelper.CalcYearStemName(user.BirthYear.Value);
                }
            }

            // JWT 沒本命星 → 用 LINE userId
            if (natalStar == 0 && !string.IsNullOrEmpty(lineUserId))
            {
                var lineUser = await _context.LineUsers.FirstOrDefaultAsync(u => u.LineUserId == lineUserId);
                if (lineUser != null) natalStar = lineUser.NatalStar;
            }

            if (natalStar == 0) return BadRequest("找不到本命星資料，請先登記生辰");

            var now = DateTime.Now;
            int dayStar   = NsCalcDayStar(now);
            int monthStar = NsCalcMonthStar(now);
            int yearStar  = NsCalcYearStar(now.Year);
            int hourStar  = NsCalcHourStar(now);
            int currentYun = Ecanapi.Services.NineStarCalcHelper.GetCurrentYun(now.Year);
            var (yunLabel, isProspering) = Ecanapi.Services.NineStarCalcHelper.GetStarYunStatus(natalStar, currentYun);

            // 查 KB（本命星 × 流日星）—— 保留原 Gemini 生成邏輯
            var rule = await _context.NineStarDailyRules
                .FirstOrDefaultAsync(r => r.NatalStar == natalStar && r.FlowStar == dayStar);

            if (rule == null || string.IsNullOrEmpty(rule.FortuneText))
            {
                var generated = await NsGeminiGenDailyAdvice(natalStar, dayStar);
                if (rule == null)
                {
                    rule = new NineStarDailyRule { NatalStar = natalStar, FlowStar = dayStar };
                    _context.NineStarDailyRules.Add(rule);
                }
                rule.FortuneText = generated.FortuneText;
                rule.Auspicious = generated.Auspicious;
                rule.Avoid = generated.Avoid;
                rule.Direction = generated.Direction;
                rule.Color = generated.Color;
                rule.UpdatedAt = DateTime.UtcNow;
                await _context.SaveChangesAsync();
                _logger.LogInformation("NineStar daily rule {Natal}x{Day} generated by Gemini and saved to KB", natalStar, dayStar);
            }

            // 九宮飛星五層組合分析
            var combinationRules = await _context.NineStarCombinationRules.ToListAsync();

            // 流年天干（供十神推算）
            string yearStemName = Ecanapi.Services.NineStarCalcHelper.CalcYearStemName(now.Year);

            // 流年暗劍殺判定：找年星布盤後五黃所在宮位，其對宮即為暗劍殺
            int? darkSwordPalace = null;
            for (int p = 1; p <= 9; p++)
            {
                if (Ecanapi.Services.NineStarCalcHelper.FlyingStarForPalace(yearStar, p, false) == 5)
                {
                    darkSwordPalace = p switch { 1 => 9, 9 => 1, 2 => 8, 8 => 2, 3 => 7, 7 => 3, 4 => 6, 6 => 4, _ => (int?)null };
                    break;
                }
            }

            // 陽遁 / 陰遁（決定流日布宮方向）
            bool dayIsForward = Ecanapi.Services.NineStarCalcHelper.GetDayIsForward(now);

            // 五層：(名稱, 中宮星, 逆飛?, 是否流年層)
            // 年：順飛（農民曆確認）
            // 月：順飛（農民曆確認）
            // 日：陽遁=順飛, 陰遁=逆飛
            // 時：一律順飛（入中序由 CalcHourStar 陽/陰遁處理）
            var layers = new[]
            {
                (pair: "三元九運", center: currentYun, reverse: false,         isYear: false),
                (pair: "流年",     center: yearStar,   reverse: false,         isYear: true),
                (pair: "流月",     center: monthStar,  reverse: false,         isYear: false),
                (pair: "流日",     center: dayStar,    reverse: !dayIsForward, isYear: false),
                (pair: "流時",     center: hourStar,   reverse: false,         isYear: false),
            };

            var combinations = layers.Select(c =>
            {
                // 計算飛入本命宮位的客星
                int visitingStar = Ecanapi.Services.NineStarCalcHelper.FlyingStarForPalace(c.center, natalStar, c.reverse);

                // 流年特殊：五黃入宅、暗劍殺
                string? specialOverride = null;
                string? tenGodText = null;
                if (c.isYear)
                {
                    if (visitingStar == 5)
                        specialOverride = "五黃入本命宮，主大凶，慎防意外病災";
                    else if (darkSwordPalace.HasValue && natalStar == darkSwordPalace.Value)
                        specialOverride = "本命宮逢暗劍殺，主凶，需謹慎行事";

                    if (!string.IsNullOrEmpty(birthStem))
                    {
                        var (tg, theme) = Ecanapi.Services.NineStarCalcHelper.CalcTenGod(birthStem, yearStemName);
                        if (!string.IsNullOrEmpty(tg))
                            tenGodText = $"{tg}年（{theme}）";
                    }
                }

                // 查33則KB（本命星=StarA，飛入客星=StarB）
                var named = combinationRules.FirstOrDefault(r =>
                    r.StarA == natalStar && r.StarB == visitingStar &&
                    (r.IsProspering == null || r.IsProspering == isProspering));

                string title, verdict, description;
                if (specialOverride != null)
                {
                    title = c.isYear && visitingStar == 5 ? "五黃大凶" : "暗劍殺";
                    verdict = c.isYear && visitingStar == 5 ? "大凶" : "凶";
                    description = specialOverride;
                }
                else if (named != null)
                {
                    title = named.Title;
                    verdict = named.Verdict;
                    description = named.Description;
                }
                else
                {
                    // fallback：星宮五行關係
                    var palaceRel = Ecanapi.Services.NineStarCalcHelper.GetStarPalaceRelation(visitingStar, natalStar);
                    var (fRelation, fVerdict, fNote) = Ecanapi.Services.NineStarCalcHelper.CalcFiveElementCombination(natalStar, visitingStar);
                    title = string.IsNullOrEmpty(palaceRel) ? fRelation : palaceRel;
                    verdict = fVerdict;
                    description = fNote;
                }

                string modified = verdict == "特殊"
                    ? (isProspering ? "最吉（同星得運）" : "最凶（同星失運）")
                    : (specialOverride != null ? verdict
                    : Ecanapi.Services.NineStarCalcHelper.ApplyYunModifier(verdict, isProspering));

                return new
                {
                    pair = c.pair,
                    starA = new { number = natalStar, name = StarNames[natalStar] },
                    starB = new { number = visitingStar, name = StarNames[visitingStar] },
                    title,
                    verdict,
                    modified,
                    description,
                    tenGodSupplement = tenGodText
                };
            }).ToList();

            // 整體評語（三元九運>流年>流月>流日）
            string overallVerdict = $"本命星目前{yunLabel}，三元{combinations[0].modified}，今年{combinations[1].modified}，本月{combinations[2].modified}，今日{combinations[3].modified}。";

            return Ok(new
            {
                date = now.ToString("yyyy-MM-dd"),
                natalStar = new { number = natalStar, name = StarNames[natalStar] },
                yunStatus = yunLabel,
                isProspering,
                yearStar  = new { number = yearStar,  name = StarNames[yearStar]  },
                monthStar = new { number = monthStar, name = StarNames[monthStar] },
                dayStar   = new { number = dayStar,   name = StarNames[dayStar]   },
                hourStar  = new { number = hourStar,  name = StarNames[hourStar]  },
                combinations,
                overallVerdict,
                fortuneText = rule.FortuneText,
                auspicious  = rule.Auspicious,
                avoid       = rule.Avoid,
                direction   = rule.Direction ?? StarDirections[natalStar],
                color       = rule.Color ?? StarColors[natalStar]
            });
        }

        // ============================================================
        //  計算層（純 C#，不呼叫 AI）
        // ============================================================

        /// <summary>計算本命星（考慮立春 2/4 前出生使用前一年）</summary>
        private static int NsCalcNatalStar(int year, int month, int day, string gender)
        {
            // 立春前（約2月4日）出生，使用前一年
            int y = (month < 2 || (month == 2 && day < 4)) ? year - 1 : year;
            return NsCalcYearStar(y, gender);
        }

        /// <summary>計算指定年份的入中星（陽遁/男=年飛星；陰遁/女=對宮反算）</summary>
        // 統一委派給 NineStarCalcHelper，避免公式重複或不一致
        private static int NsCalcYearStar(int year, string gender = "M")
            => Ecanapi.Services.NineStarCalcHelper.CalcYearStar(year, gender == "F" ? 2 : 1);

        /// <summary>計算流年星（入中宮，通用）</summary>
        private static int NsCalcYearStar(int year) => NsCalcYearStar(year, "M");

        /// <summary>計算流月星（依節氣月）</summary>
        private static int NsCalcMonthStar(DateTime date)
        {
            int yearStar = NsCalcYearStar(date.Year);
            int solarMonth = NsGetSolarMonth(date); // 1=寅月(立春後)

            // 月首星：年星 1/4/7 → 寅月起8；2/5/8 → 寅月起5；3/6/9 → 寅月起2
            int[] monthStartStars = { 0, 8, 5, 2, 8, 5, 2, 8, 5, 2 };
            int startStar = monthStartStars[yearStar];

            // 每月遞減（月星逆排）
            int star = ((startStar - solarMonth + 1) % 9 + 9) % 9;
            return star == 0 ? 9 : star;
        }

        /// <summary>計算流日星（依節氣區間 + 甲子日順逆行）</summary>
        private static int NsCalcDayStar(DateTime date)
        {
            // 六個節氣區間與起始星、方向
            // 冬至~雨水：起1順；雨水~穀雨：起7順；穀雨~夏至：起4順
            // 夏至~處暑：起9逆；處暑~霜降：起3逆；霜降~冬至：起6逆
            var (startStar, forward, periodStart) = NsGetDayStarPeriod(date);

            // 找本區間內第一個甲子日（從 periodStart 往前找最近甲子日）
            DateTime jiazi = NsGetLastJiaZiDay(periodStart);

            int daysDiff = (int)(date.Date - jiazi).TotalDays;

            int star;
            if (forward)
                star = ((startStar - 1 + daysDiff) % 9 + 9) % 9 + 1;
            else
                star = ((startStar - 1 - daysDiff) % 9 + 9) % 9 + 1;

            return star;
        }

        /// <summary>計算流時星（依日支類型 + 半年陰陽）</summary>
        private static int NsCalcHourStar(DateTime dateTime)
        {
            // 取時辰序號（子=0, 丑=1, ...亥=11），每2小時一個時辰
            int hour = dateTime.Hour;
            int branchIdx = hour / 2; // 0=子,1=丑,...,11=亥（23:00-00:59 = 子）
            if (hour == 23) branchIdx = 0;

            // 日支（60甲子循環的地支）
            int cycle = NsGet60CycleIndex(dateTime);
            int dayBranch = cycle % 12; // 甲子=0→子,乙丑=1→丑,...

            // 孟(寅申巳亥)/仲(子午卯酉)/季(辰戌丑未)
            int[] mengBranches = { 2, 5, 8, 11 }; // 寅巳申亥
            int[] zhongBranches = { 0, 3, 6, 9 }; // 子卯午酉
            // 其餘為季：辰戌丑未

            int startStar;
            bool isYangHalf = NsIsYangHalf(dateTime); // 冬至到夏至為陽
            if (mengBranches.Contains(dayBranch))
                startStar = isYangHalf ? 7 : 3;
            else if (zhongBranches.Contains(dayBranch))
                startStar = isYangHalf ? 1 : 9;
            else
                startStar = isYangHalf ? 4 : 6;

            // 時辰從子時順行
            int star = ((startStar - 1 + branchIdx) % 9 + 9) % 9 + 1;
            return star;
        }

        // ============================================================
        //  輔助計算
        // ============================================================

        /// <summary>取日期在60甲子中的 index（0=甲子...59=癸亥）</summary>
        private static int NsGet60CycleIndex(DateTime date)
        {
            int days = (int)(date.Date - NsEpochDate).TotalDays;
            return ((NsEpochCycleIndex + days) % 60 + 60) % 60;
        }

        /// <summary>取指定日期之前（含當天）最近的甲子日</summary>
        private static DateTime NsGetLastJiaZiDay(DateTime from)
        {
            int idx = NsGet60CycleIndex(from);
            int daysBack = idx; // 甲子=index 0，往前 idx 天即到甲子
            return from.Date.AddDays(-daysBack);
        }

        /// <summary>判斷日期屬於哪個日飛星區間，回傳（起始星, 順行, 區間起始日）</summary>
        private static (int startStar, bool forward, DateTime periodStart) NsGetDayStarPeriod(DateTime date)
        {
            int year = date.Year;
            int doy = date.DayOfYear;

            // 六個節氣的近似日（月/日）→ 換算為 DOY
            // 冬至 ~Dec22, 雨水 ~Feb19, 穀雨 ~Apr20, 夏至 ~Jun21, 處暑 ~Aug23, 霜降 ~Oct23
            // 使用前一年冬至到今年冬至橫跨兩年，需特別處理
            int doyYuShui = 50;  // 雨水 ~Feb19
            int doyGuYu = 110;   // 穀雨 ~Apr20
            int doyXiaZhi = 172; // 夏至 ~Jun21
            int doyChuShu = 235; // 處暑 ~Aug23
            int doyShuangJiang = 296; // 霜降 ~Oct23
            // 冬至：12月22日 = DOY 356(平年)/357(閏年)
            int doyDongZhi = DateTime.IsLeapYear(year) ? 357 : 356;

            if (doy < doyYuShui)
            {
                // 上年冬至 ~ 雨水：起1，順行
                // 區間起始 = 上年冬至 (約 Dec 22 of year-1)
                var prevDongZhi = new DateTime(year - 1, 12, 22);
                return (1, true, prevDongZhi);
            }
            else if (doy < doyGuYu)
            {
                // 雨水 ~ 穀雨：起7，順行
                return (7, true, new DateTime(year, 2, 19));
            }
            else if (doy < doyXiaZhi)
            {
                // 穀雨 ~ 夏至：起4，順行
                return (4, true, new DateTime(year, 4, 20));
            }
            else if (doy < doyChuShu)
            {
                // 夏至 ~ 處暑：起9，逆行
                return (9, false, new DateTime(year, 6, 21));
            }
            else if (doy < doyShuangJiang)
            {
                // 處暑 ~ 霜降：起3，逆行
                return (3, false, new DateTime(year, 8, 23));
            }
            else if (doy < doyDongZhi)
            {
                // 霜降 ~ 冬至：起6，逆行
                return (6, false, new DateTime(year, 10, 23));
            }
            else
            {
                // 冬至後 ~ 年底（下一個雨水前）：起1，順行
                return (1, true, new DateTime(year, 12, 22));
            }
        }

        /// <summary>判斷是否在陽遁半年（冬至 ~ 夏至）</summary>
        private static bool NsIsYangHalf(DateTime date)
        {
            int doy = date.DayOfYear;
            int doyXiaZhi = DateTime.IsLeapYear(date.Year) ? 173 : 172;
            int doyDongZhi = DateTime.IsLeapYear(date.Year) ? 357 : 356;
            // 陽遁：冬至後（DOY >= doyDongZhi 或 DOY < doyXiaZhi）
            return doy >= doyDongZhi || doy < doyXiaZhi;
        }

        /// <summary>取節氣月份（1=寅月/立春後，12=丑月/小寒後）</summary>
        private static int NsGetSolarMonth(DateTime date)
        {
            int y = date.Year;
            int m = date.Month;
            int d = date.Day;

            // 12個節氣起始日（月, 日）對應節月 1-12（1=寅月）
            // 小寒 Jan6=月12，立春 Feb4=月1，驚蟄 Mar6=月2，清明 Apr5=月3
            // 立夏 May6=月4，芒種 Jun6=月5，小暑 Jul7=月6，立秋 Aug7=月7
            // 白露 Sep8=月8，寒露 Oct8=月9，立冬 Nov7=月10，大雪 Dec7=月11
            (int sm, int sd, int solarM)[] terms =
            {
                (1,  6,  12), // 小寒 → 丑月
                (2,  4,  1),  // 立春 → 寅月
                (3,  6,  2),  // 驚蟄 → 卯月
                (4,  5,  3),  // 清明 → 辰月
                (5,  6,  4),  // 立夏 → 巳月
                (6,  6,  5),  // 芒種 → 午月
                (7,  7,  6),  // 小暑 → 未月
                (8,  7,  7),  // 立秋 → 申月
                (9,  8,  8),  // 白露 → 酉月
                (10, 8,  9),  // 寒露 → 戌月
                (11, 7,  10), // 立冬 → 亥月
                (12, 7,  11), // 大雪 → 子月
            };

            int current = 11; // 預設子月
            foreach (var (sm, sd, solarM) in terms)
            {
                if (m > sm || (m == sm && d >= sd))
                    current = solarM;
            }
            return current;
        }

        // ============================================================
        //  Gemini 生成（暫代 KB，結果自動存回 DB）
        // ============================================================

        private async Task<string> NsGeminiGenTrait(int starNumber, string type)
        {
            string starName = StarNames[starNumber];
            string typeName = type switch
            {
                "personality" => "個性特質與人生格局",
                "career" => "事業財運傾向",
                "relationship" => "感情與人際關係",
                "health" => "健康養生建議",
                _ => "整體特質"
            };

            string prompt = $"請以九星氣學的角度，為【{starName}】本命星的人撰寫「{typeName}」說明。" +
                            $"約200字，繁體中文，語氣直接對命主說話，不要出現「九星氣學」等學術名詞，" +
                            $"只給論斷與建議。";

            return await NsCallGemini(prompt);
        }

        private async Task<(string FortuneText, string Auspicious, string Avoid, string Direction, string Color)>
            NsGeminiGenDailyAdvice(int natalStar, int dayStar)
        {
            string prompt = $"九星氣學今日建議：本命星【{StarNames[natalStar]}】，今日流日星【{StarNames[dayStar]}】。" +
                            $"請以繁體中文輸出 JSON，包含以下欄位：" +
                            $"fortune_text（今日整體運勢，約100字），" +
                            $"auspicious（今日宜做的事，20字內），" +
                            $"avoid（今日不宜，20字內），" +
                            $"direction（今日吉方位，10字內），" +
                            $"color（今日吉顏色，10字內）。" +
                            $"直接回傳 JSON，不要 markdown。";

            try
            {
                string raw = await NsCallGemini(prompt);
                var json = JsonSerializer.Deserialize<JsonElement>(raw);
                return (
                    json.GetProperty("fortune_text").GetString() ?? "",
                    json.GetProperty("auspicious").GetString() ?? "",
                    json.GetProperty("avoid").GetString() ?? "",
                    json.GetProperty("direction").GetString() ?? StarDirections[natalStar],
                    json.GetProperty("color").GetString() ?? StarColors[natalStar]
                );
            }
            catch
            {
                // JSON 解析失敗，回傳純文字版本
                string raw = await NsCallGemini(
                    $"九星氣學：本命星{StarNames[natalStar]}，今日流星{StarNames[dayStar]}，" +
                    $"請給一段今日建議（100字，繁體中文）。");
                return (raw, "", "", StarDirections[natalStar], StarColors[natalStar]);
            }
        }

        private async Task<string> NsCallGemini(string prompt)
        {
            var apiKey = _config["Gemini:ApiKey"];
            var url = $"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={apiKey}";
            var payload = new { contents = new[] { new { parts = new[] { new { text = prompt } } } } };
            var response = await _httpClient.PostAsJsonAsync(url, payload);
            var rawJson = await response.Content.ReadAsStringAsync();
            var json = JsonSerializer.Deserialize<JsonElement>(rawJson);
            if (!json.TryGetProperty("candidates", out var candidates))
                throw new Exception($"Gemini 回傳錯誤: {rawJson[..Math.Min(200, rawJson.Length)]}");
            return candidates[0].GetProperty("content").GetProperty("parts")[0].GetProperty("text").GetString()!;
        }

        // ============================================================
        //  Admin 手動測試 Instagram 發文
        // ============================================================

        /// <summary>Admin 手動觸發 Instagram 發文（立即執行，用於測試）</summary>
        [HttpPost("ig-post-now")]
        [Authorize]
        public async Task<IActionResult> IgPostNow()
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            var user = await _context.Users.FindAsync(userId);
            var adminEmail = _config["Admin:Email"];
            if (user == null || !string.Equals(user.Email, adminEmail, StringComparison.OrdinalIgnoreCase))
                return Forbid();

            var igService = HttpContext.RequestServices.GetRequiredService<Ecanapi.Services.InstagramDailyPostService>();
            await igService.PostDailyFortuneAsync();
            return Ok(new { message = "Instagram 發文已觸發" });
        }
    }

    // ============================================================
    //  Request DTO
    // ============================================================
    public class LineRegisterRequest
    {
        public string LineUserId { get; set; } = string.Empty;
        public int BirthYear { get; set; }
        public int BirthMonth { get; set; }
        public int BirthDay { get; set; }
        public string Gender { get; set; } = string.Empty; // M / F
        public string? DisplayName { get; set; }
    }

    public class SetNotifyRequest
    {
        public bool Enabled { get; set; }
    }
}
