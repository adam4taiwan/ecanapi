using Ecanapi.Data;
using Ecanapi.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Security.Claims;
using System.Text;
using System.Text.Json;

namespace Ecanapi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class FortuneController : ControllerBase
    {
        private readonly ApplicationDbContext _context;
        private readonly IConfiguration _config;
        private readonly ILogger<FortuneController> _logger;
        private static readonly HttpClient _httpClient = new HttpClient { Timeout = TimeSpan.FromMinutes(3) };

        // 天干地支循環
        private static readonly string[] TianGan = { "甲", "乙", "丙", "丁", "戊", "己", "庚", "辛", "壬", "癸" };
        private static readonly string[] DiZhi = { "子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥" };

        public FortuneController(ApplicationDbContext context, IConfiguration config, ILogger<FortuneController> logger)
        {
            _context = context;
            _config = config;
            _logger = logger;
        }

        /// <summary>取得今日個人運勢（每日免費，自動快取）</summary>
        [HttpGet("daily")]
        [Authorize]
        public async Task<IActionResult> GetDailyFortune()
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (string.IsNullOrEmpty(userId))
                return Unauthorized(new { error = "請重新登入" });

            var user = await _context.Users.FindAsync(userId);
            if (user == null) return Unauthorized(new { error = "找不到用戶" });

            var today = DateTime.UtcNow.Date;

            // 檢查今日是否已有快取
            var cached = await _context.DailyFortunes
                .FirstOrDefaultAsync(f => f.UserId == userId && f.FortuneDate == today);

            if (cached != null)
            {
                return Ok(new { content = cached.Content, date = FormatChineseDate(today), cached = true });
            }

            // 產生今日干支
            string todayGanZhi = GetGanZhi(today);
            string todayDateStr = today.ToString("yyyy年M月d日");

            // 建立 Prompt
            string prompt;
            if (user.HasBirthData)
            {
                int birthYear = user.BirthYear!.Value;
                int birthMonth = user.BirthMonth!.Value;
                int birthDay = user.BirthDay!.Value;
                int birthHour = user.BirthHour!.Value;
                int gender = user.BirthGender ?? 1;
                string genderText = gender == 1 ? "男（乾造）" : "女（坤造）";
                int age = today.Year - birthYear;

                prompt = $@"你是命理鑑定大師『玉洞子』。請根據命主生辰與今日干支，給出簡潔有力的今日運勢指引。

【命主資料】
生辰：{birthYear}年{birthMonth}月{birthDay}日 {birthHour}時
性別：{genderText}
今年歲數：約 {age} 歲

【今日資訊】
日期：{todayDateStr}
日柱干支：{todayGanZhi}

【輸出格式（嚴格遵守，禁止多寫）】

=== {todayDateStr} 今日運勢 ===

【總評】一句話點出今日整體氣場（15字以內）

【財運】具體斷語，20字以內，直接說吉或凶及原因
【事業】具體斷語，20字以內，直接說吉或凶及原因
【感情】具體斷語，20字以內，直接說吉或凶及原因
【健康】具體斷語，20字以內，點出需注意部位

【今日提醒】一句可執行的建議（20字以內）

-----------------------------------------------------------------
鑑定大師：玉洞子  |  知日善用，趨吉避凶。";
            }
            else
            {
                // 無生辰資料，給通用今日運勢
                prompt = $@"你是命理鑑定大師『玉洞子』。請根據今日干支，給出今日的通用運勢指引。

【今日資訊】
日期：{todayDateStr}
日柱干支：{todayGanZhi}

【輸出格式（嚴格遵守，禁止多寫）】

=== {todayDateStr} 今日運勢 ===

【總評】一句話點出今日干支整體氣場（15字以內）

【財運】今日財運斷語，20字以內
【事業】今日事業斷語，20字以內
【感情】今日感情斷語，20字以內
【健康】今日健康提示，20字以內

【今日提醒】一句通用建議（20字以內）

-----------------------------------------------------------------
鑑定大師：玉洞子  |  知日善用，趨吉避凶。";
            }

            try
            {
                string content = await CallGeminiApi(prompt);

                // 儲存快取
                var fortune = new DailyFortune
                {
                    UserId = userId,
                    FortuneDate = today,
                    Content = content,
                    CreatedAt = DateTime.UtcNow
                };
                _context.DailyFortunes.Add(fortune);
                await _context.SaveChangesAsync();

                return Ok(new { content, date = FormatChineseDate(today), cached = false });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "每日運勢生成失敗 UserId={UserId}", userId);
                return StatusCode(500, new { error = "今日運勢生成失敗，請稍後再試" });
            }
        }

        /// <summary>取得今日個人運勢（純知識庫版，不呼叫 Gemini）</summary>
        [HttpGet("daily-kb")]
        [Authorize]
        public async Task<IActionResult> GetDailyFortuneKb()
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (string.IsNullOrEmpty(userId))
                return Unauthorized(new { error = "請重新登入" });

            var today = DateTime.UtcNow.Date;

            // 檢查今日是否已有快取
            var cached = await _context.DailyFortunes
                .FirstOrDefaultAsync(f => f.UserId == userId && f.FortuneDate == today);

            if (cached != null)
                return Ok(new { content = cached.Content, date = FormatChineseDate(today), cached = true });

            // 計算今日日柱干支
            string ganZhi = GetGanZhi(today);
            string tianGan = ganZhi[..1];   // e.g. 壬
            string diZhi   = ganZhi[1..];   // e.g. 子

            // 查詢知識庫規則
            var ganRules = await _context.FortuneRules
                .Where(r => r.Subcategory == "日干" && r.Title == tianGan && r.IsActive)
                .ToListAsync();

            var zhiRules = await _context.FortuneRules
                .Where(r => r.Subcategory == "日支" && r.Title == diZhi && r.IsActive)
                .ToListAsync();

            if (!ganRules.Any() && !zhiRules.Any())
            {
                _logger.LogWarning("daily-kb: 知識庫無日柱規則 GanZhi={GanZhi}", ganZhi);
                return StatusCode(503, new { error = "知識庫尚未建立，請先在後台匯入日柱規則" });
            }

            // 按 ConditionText 建索引方便取用
            var gan = ganRules.ToDictionary(r => r.ConditionText ?? "", r => r.ResultText);
            var zhi = zhiRules.ToDictionary(r => r.ConditionText ?? "", r => r.ResultText);

            string todayDateStr = today.ToString("yyyy年M月d日");
            var sb = new StringBuilder();

            sb.AppendLine($"=== {todayDateStr} 今日運勢 ===");
            sb.AppendLine();

            // 總評：日干為主，日支補充
            var summary = new StringBuilder();
            if (gan.TryGetValue("總評", out var ganSummary)) summary.Append(ganSummary);
            if (zhi.TryGetValue("總評", out var zhiSummary) && zhiSummary != ganSummary)
            {
                if (summary.Length > 0) summary.Append("，");
                summary.Append(zhiSummary);
            }
            if (summary.Length > 0)
                sb.AppendLine($"【總評】{summary}");
            sb.AppendLine();

            // 財運/事業/感情/健康：以日干規則為主
            foreach (var domain in new[] { "財運", "事業", "感情", "健康" })
            {
                if (gan.TryGetValue(domain, out var val))
                    sb.AppendLine($"【{domain}】{val}");
            }
            sb.AppendLine();

            // 今日提醒：以日支規則為主（日干補充）
            var reminder = zhi.TryGetValue("提醒", out var zhiReminder) ? zhiReminder
                         : gan.TryGetValue("提醒", out var ganReminder) ? ganReminder
                         : null;
            if (reminder != null)
                sb.AppendLine($"【今日提醒】{reminder}");

            sb.AppendLine();
            sb.AppendLine("-----------------------------------------------------------------");
            sb.AppendLine($"鑑定大師：玉洞子  |  知日善用，趨吉避凶。");

            string content = sb.ToString().TrimEnd();

            // 儲存快取
            _context.DailyFortunes.Add(new DailyFortune
            {
                UserId = userId,
                FortuneDate = today,
                Content = content,
                CreatedAt = DateTime.UtcNow
            });
            await _context.SaveChangesAsync();

            return Ok(new { content, date = FormatChineseDate(today), cached = false });
        }

        /// <summary>取得今日個人化運勢（依命主日主計算十神，純知識庫版）</summary>
        [HttpGet("daily-personal")]
        [Authorize]
        public async Task<IActionResult> GetDailyFortunePersonal()
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (string.IsNullOrEmpty(userId))
                return Unauthorized(new { error = "請重新登入" });

            var user = await _context.Users.FindAsync(userId);
            if (user == null) return Unauthorized(new { error = "找不到用戶" });

            if (!user.HasBirthData)
                return BadRequest(new { error = "請先填寫生辰資料，才能取得個人化運勢" });

            var today = DateTime.UtcNow.Date;

            // 檢查今日是否已有快取（personal 版用不同 key 避免與 daily-kb 快取衝突）
            var cacheKey = $"personal:{userId}";
            var cached = await _context.DailyFortunes
                .FirstOrDefaultAsync(f => f.UserId == cacheKey && f.FortuneDate == today);

            if (cached != null)
                return Ok(new { content = cached.Content, date = FormatChineseDate(today), cached = true });

            // 計算日主（由生辰日柱天干決定）
            var birthDate = new DateTime(user.BirthYear!.Value, user.BirthMonth!.Value, user.BirthDay!.Value);
            string birthGanZhi = GetGanZhi(birthDate);
            string riZhu = birthGanZhi[..1]; // 日主天干，e.g. 辛

            // 計算今日日柱
            string todayGanZhi = GetGanZhi(today);
            string todayTianGan = todayGanZhi[..1];
            string todayDiZhi   = todayGanZhi[1..];

            // 計算今日天干對日主的十神關係
            string shiShen = GetShiShen(riZhu, todayTianGan);

            // 查詢今日通用日柱規則（Phase 1）
            var ganRules = await _context.FortuneRules
                .Where(r => r.Subcategory == "日干" && r.Title == todayTianGan && r.IsActive)
                .ToListAsync();
            var zhiRules = await _context.FortuneRules
                .Where(r => r.Subcategory == "日支" && r.Title == todayDiZhi && r.IsActive)
                .ToListAsync();

            // 查詢十神應事規則（Phase 2 個人加成）
            var shiShenRules = await _context.FortuneRules
                .Where(r => r.Subcategory == "十神應事" && r.Title == shiShen && r.IsActive)
                .ToListAsync();

            string todayDateStr = today.ToString("yyyy年M月d日");
            var sb = new StringBuilder();

            sb.AppendLine($"=== {todayDateStr} 今日運勢 ===");
            sb.AppendLine();

            // 通用日柱總評
            var gan = ganRules.ToDictionary(r => r.ConditionText ?? "", r => r.ResultText);
            var zhi = zhiRules.ToDictionary(r => r.ConditionText ?? "", r => r.ResultText);

            var summary = new StringBuilder();
            if (gan.TryGetValue("總評", out var ganSummary)) summary.Append(ganSummary);
            if (zhi.TryGetValue("總評", out var zhiSummary) && zhiSummary != ganSummary)
            {
                if (summary.Length > 0) summary.Append("，");
                summary.Append(zhiSummary);
            }
            if (summary.Length > 0) sb.AppendLine($"【總評】{summary}");
            sb.AppendLine();

            // 財運/事業/感情/健康
            foreach (var domain in new[] { "財運", "事業", "感情", "健康" })
                if (gan.TryGetValue(domain, out var val)) sb.AppendLine($"【{domain}】{val}");

            sb.AppendLine();

            // 今日提醒
            var reminder = zhi.TryGetValue("提醒", out var zr) ? zr
                         : gan.TryGetValue("提醒", out var gr) ? gr : null;
            if (reminder != null) sb.AppendLine($"【今日提醒】{reminder}");

            // === 個人命主加成（十神）===
            if (shiShenRules.Any())
            {
                sb.AppendLine();
                sb.AppendLine($"--- 命主加成（日主 {riZhu}，今日十神：{shiShen}）---");

                var good = shiShenRules.FirstOrDefault(r => r.ConditionText == "喜用")?.ResultText;
                var bad  = shiShenRules.FirstOrDefault(r => r.ConditionText == "忌用")?.ResultText;

                if (good != null) sb.AppendLine($"【順勢】{good}");
                if (bad  != null) sb.AppendLine($"【逆勢】{bad}");
            }

            // === Phase 3：紫微命宮主星加成 ===
            var userChart = await _context.UserCharts
                .FirstOrDefaultAsync(c => c.UserId == userId);

            if (userChart != null && !string.IsNullOrEmpty(userChart.MingGongMainStars))
            {
                var mingGongStars = userChart.MingGongMainStars.Split(',', StringSplitOptions.RemoveEmptyEntries);
                var ziweiRules = new List<(string star, string text)>();

                foreach (var star in mingGongStars)
                {
                    var rule = await _context.FortuneRules
                        .FirstOrDefaultAsync(r => r.Subcategory == "命宮主星" && r.Title == star.Trim() && r.IsActive);
                    if (rule != null)
                        ziweiRules.Add((star.Trim(), rule.ResultText));
                }

                if (ziweiRules.Any())
                {
                    sb.AppendLine();
                    sb.AppendLine($"--- 紫微命宮加成（命宮主星：{userChart.MingGongMainStars}）---");
                    foreach (var (star, text) in ziweiRules)
                        sb.AppendLine($"【{star}】{text}");
                }
            }

            sb.AppendLine();
            sb.AppendLine("-----------------------------------------------------------------");
            sb.AppendLine($"鑑定大師：玉洞子  |  知日善用，趨吉避凶。");

            string content = sb.ToString().TrimEnd();

            // 儲存快取
            _context.DailyFortunes.Add(new DailyFortune
            {
                UserId = cacheKey,
                FortuneDate = today,
                Content = content,
                CreatedAt = DateTime.UtcNow
            });
            await _context.SaveChangesAsync();

            return Ok(new { content, date = FormatChineseDate(today), cached = false, riZhu, shiShen });
        }

        /// <summary>計算天干甲相對日主的十神關係</summary>
        private static string GetShiShen(string dayMaster, string stem)
        {
            if (dayMaster == stem) return "比肩";

            var element = new Dictionary<string, string>
            {
                {"甲","木"},{"乙","木"},{"丙","火"},{"丁","火"},
                {"戊","土"},{"己","土"},{"庚","金"},{"辛","金"},
                {"壬","水"},{"癸","水"}
            };
            var isYang = new Dictionary<string, bool>
            {
                {"甲",true},{"乙",false},{"丙",true},{"丁",false},
                {"戊",true},{"己",false},{"庚",true},{"辛",false},
                {"壬",true},{"癸",false}
            };
            var generates = new Dictionary<string, string>
                { {"木","火"},{"火","土"},{"土","金"},{"金","水"},{"水","木"} };
            var controls = new Dictionary<string, string>
                { {"木","土"},{"土","水"},{"水","火"},{"火","金"},{"金","木"} };

            string dmElem = element[dayMaster];
            string stElem = element[stem];
            bool sameYY   = isYang[dayMaster] == isYang[stem];

            if (dmElem == stElem)
                return sameYY ? "比肩" : "劫財";

            if (generates.TryGetValue(dmElem, out var genTarget) && genTarget == stElem)
                return sameYY ? "食神" : "傷官";

            if (controls.TryGetValue(dmElem, out var ctrlTarget) && ctrlTarget == stElem)
                return sameYY ? "偏財" : "正財";

            if (controls.TryGetValue(stElem, out var ctrlMe) && ctrlMe == dmElem)
                return sameYY ? "七殺" : "正官";

            if (generates.TryGetValue(stElem, out var genMe) && genMe == dmElem)
                return sameYY ? "偏印" : "正印";

            return "比肩"; // fallback
        }

        private string GetGanZhi(DateTime date)
        {
            // 以1900年1月31日（甲子日）為基準計算日柱干支
            var baseDate = new DateTime(1900, 1, 31);
            int days = (int)(date - baseDate).TotalDays;
            int ganIndex = ((days % 10) + 10) % 10;
            int zhiIndex = ((days % 12) + 12) % 12;
            return TianGan[ganIndex] + DiZhi[zhiIndex];
        }

        private string FormatChineseDate(DateTime date)
        {
            return $"{date.Year}年{date.Month}月{date.Day}日";
        }

        private async Task<string> CallGeminiApi(string prompt)
        {
            var apiKey = _config["Gemini:ApiKey"];
            var url = $"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={apiKey}";
            var payload = new { contents = new[] { new { parts = new[] { new { text = prompt } } } } };
            var response = await _httpClient.PostAsJsonAsync(url, payload);
            var rawJson = await response.Content.ReadAsStringAsync();
            _logger.LogInformation("Gemini Fortune {StatusCode}: {Preview}", (int)response.StatusCode,
                rawJson.Length > 200 ? rawJson[..200] : rawJson);
            var json = JsonSerializer.Deserialize<JsonElement>(rawJson);
            if (!json.TryGetProperty("candidates", out var candidates))
                throw new Exception($"Gemini 回傳錯誤: {rawJson[..Math.Min(200, rawJson.Length)]}");
            return candidates[0].GetProperty("content").GetProperty("parts")[0].GetProperty("text").GetString()!;
        }
    }
}
