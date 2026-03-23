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

        /// <summary>取得今日個人運勢（有命盤→KB個人化，無命盤→Gemini）</summary>
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

            // 有命盤且有生辰 → 導向 KB 個人化版（不呼叫 Gemini）
            if (user.HasBirthData)
            {
                var chartExists = await _context.UserCharts.AnyAsync(c => c.UserId == userId);
                if (chartExists)
                    return await GetDailyFortunePersonal();
            }

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
                bool geminiSkipGanQing = age < 16 || age >= 75;
                string geminiGanQingLine = geminiSkipGanQing ? "" : "【感情】具體斷語，20字以內，直接說吉或凶及原因\n";
                string geminiAgeNote = GetGeminiAgeNote(age);

                prompt = $@"你是命理鑑定大師『玉洞子』。請根據命主生辰與今日干支，給出簡潔有力的今日運勢指引。

【命主資料】
生辰：{birthYear}年{birthMonth}月{birthDay}日 {birthHour}時
性別：{genderText}
今年歲數：約 {age} 歲
{geminiAgeNote}
【今日資訊】
日期：{todayDateStr}
日柱干支：{todayGanZhi}

【輸出格式（嚴格遵守，禁止多寫）】

=== {todayDateStr} 今日運勢 ===

【總評】一句話點出今日整體氣場（15字以內）

【財運】具體斷語，20字以內，直接說吉或凶及原因
【事業】具體斷語，20字以內，直接說吉或凶及原因
{geminiGanQingLine}【健康】具體斷語，20字以內，點出需注意部位

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

        /// <summary>清除當前用戶今日運勢快取（更新生辰後呼叫）</summary>
        [HttpDelete("my-cache-today")]
        [Authorize]
        public async Task<IActionResult> ClearMyFortuneToday()
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (string.IsNullOrEmpty(userId))
                return Unauthorized(new { error = "請重新登入" });

            var today = DateTime.UtcNow.Date;
            var cacheKey = $"personal:{userId}";

            var records = _context.DailyFortunes.Where(f =>
                f.FortuneDate == today &&
                (f.UserId == userId || f.UserId == cacheKey));

            _context.DailyFortunes.RemoveRange(records);
            await _context.SaveChangesAsync();
            return Ok(new { cleared = true });
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

            // 載入用戶取得年齡（供年齡適切性過濾）
            var kbUser = await _context.Users.FindAsync(userId);
            int kbAge = (kbUser?.BirthYear != null) ? today.Year - kbUser.BirthYear.Value : 30;

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

            // 財運/事業/感情/健康（感情視年齡過濾）
            bool kbSkipGanQing = kbAge < 16 || kbAge >= 75;
            foreach (var domain in new[] { "財運", "事業", "感情", "健康" })
            {
                if (domain == "感情" && kbSkipGanQing) continue;
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
            int currentAge = today.Year - user.BirthYear!.Value;

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

            // 年齡適切性提示
            string ageHint = GetDailyAgeTopicHint(currentAge);
            if (!string.IsNullOrEmpty(ageHint))
            {
                sb.AppendLine(ageHint);
                sb.AppendLine();
            }

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

            // 財運/事業/感情/健康（感情視年齡過濾）
            bool skipGanQing = currentAge < 16 || currentAge >= 75;
            foreach (var domain in new[] { "財運", "事業", "感情", "健康" })
            {
                if (domain == "感情" && skipGanQing) continue;
                if (gan.TryGetValue(domain, out var val)) sb.AppendLine($"【{domain}】{val}");
            }

            sb.AppendLine();

            // 今日提醒
            var reminder = zhi.TryGetValue("提醒", out var zr) ? zr
                         : gan.TryGetValue("提醒", out var gr) ? gr : null;
            if (reminder != null) sb.AppendLine($"【今日提醒】{reminder}");

            // === 載入命盤（供 Phase 2/4/5 共用）===
            var userChart = await _context.UserCharts.FirstOrDefaultAsync(c => c.UserId == userId);

            // 從 ChartJson 一次性提取所有所需資料
            string[] chartStems    = Array.Empty<string>();   // [年干, 月干, 時干]
            string[] chartBranches = Array.Empty<string>();   // [年支, 月支, 日支, 時支]
            string chartMonthBranch = "";
            Dictionary<string, string>? chartBranchDict = null;
            string? daYunLiuShen = null, daYunStem = null, daYunBranch = null;

            if (userChart != null && !string.IsNullOrEmpty(userChart.ChartJson))
            {
                try
                {
                    using var chartDoc = JsonDocument.Parse(userChart.ChartJson);
                    var root = chartDoc.RootElement;

                    // 四柱（calculate API 回傳 "bazi"，前端直接存入 ChartJson）
                    if (root.TryGetProperty("bazi", out var baziEl) ||
                        root.TryGetProperty("baziInfo", out baziEl) ||
                        root.TryGetProperty("BaziInfo", out baziEl))
                    {
                        var (yStem, yBranch) = ExtractPillar(baziEl, "yearPillar");
                        var (mStem, mBranch) = ExtractPillar(baziEl, "monthPillar");
                        var (dStem, dBranch) = ExtractPillar(baziEl, "dayPillar");
                        var (tStem, tBranch) = ExtractPillar(baziEl, "timePillar");
                        chartStems    = new[] { yStem, mStem, tStem };
                        chartBranches = new[] { yBranch, mBranch, dBranch, tBranch };
                        chartMonthBranch = mBranch;
                        chartBranchDict  = new Dictionary<string, string>
                        {
                            {"年支", yBranch}, {"月支", mBranch}, {"日支", dBranch}, {"時支", tBranch}
                        };
                    }

                    // 大運
                    if (root.TryGetProperty("baziLuckCycles", out var luckEl) ||
                        root.TryGetProperty("BaziLuckCycles", out luckEl))
                    {
                        foreach (var cycle in luckEl.EnumerateArray())
                        {
                            int startAge = cycle.TryGetProperty("startAge", out var sa) ? sa.GetInt32() : 0;
                            int endAge   = cycle.TryGetProperty("endAge",   out var ea) ? ea.GetInt32() : 0;
                            if (currentAge >= startAge && currentAge <= endAge)
                            {
                                daYunLiuShen = cycle.TryGetProperty("liuShen",       out var ls) ? ls.GetString() : null;
                                daYunStem    = cycle.TryGetProperty("heavenlyStem",   out var hs) ? hs.GetString() : null;
                                daYunBranch  = cycle.TryGetProperty("earthlyBranch", out var eb) ? eb.GetString() : null;
                                break;
                            }
                        }
                    }
                }
                catch { /* ChartJson 解析失敗時靜默跳過 */ }
            }

            // === 個人命主加成（十神 + 身強弱）===
            if (shiShenRules.Any())
            {
                // 計算身強弱
                string bodyLabel = "";
                bool? isShun = null; // null = 中和（顯示兩者）

                if (!string.IsNullOrEmpty(chartMonthBranch))
                {
                    int score = CalcBodyStrength(riZhu, chartMonthBranch, chartStems, chartBranches);
                    if (score >= 5)
                    {
                        bodyLabel = "身強";
                        isShun = StrongYongShen.Contains(shiShen);
                    }
                    else if (score <= 2)
                    {
                        bodyLabel = "身弱";
                        isShun = WeakYongShen.Contains(shiShen);
                    }
                    else
                    {
                        bodyLabel = "中和";
                        isShun = null;
                    }
                }

                sb.AppendLine();
                string headerLabel = bodyLabel != "" ? $" | {bodyLabel}" : "";
                sb.AppendLine($"--- 命主加成（日主 {riZhu}，今日十神：{shiShen}{headerLabel}）---");

                var good = shiShenRules.FirstOrDefault(r => r.ConditionText == "喜用")?.ResultText;
                var bad  = shiShenRules.FirstOrDefault(r => r.ConditionText == "忌用")?.ResultText;

                if (isShun == null)
                {
                    // 中和：顯示兩者
                    if (good != null) sb.AppendLine($"【順勢】{good}");
                    if (bad  != null) sb.AppendLine($"【逆勢】{bad}");
                }
                else if (isShun == true)
                {
                    if (good != null) sb.AppendLine($"【順勢】{good}");
                }
                else
                {
                    if (bad != null) sb.AppendLine($"【逆勢】{bad}");
                }

                // 刑沖會合破害：今日日支 vs 命局四柱地支
                if (chartBranchDict != null && chartBranchDict.Any(kv => !string.IsNullOrEmpty(kv.Value)))
                {
                    var relations = BranchRelationModels.BaZiBranchRelation.CalcBranchRelations(
                        chartBranchDict, todayDiZhi, "今日");
                    foreach (var rel in relations)
                    {
                        string impact = GetRelationImpact(rel.RelationType, isShun);
                        string dirText = isShun == true ? "順勢" : isShun == false ? "逆勢" : "運勢";
                        sb.AppendLine($"【今日{todayDiZhi}{rel.RelationType}{rel.TargetPillar}{rel.TargetBranch}】{dirText}受{rel.RelationType}，{impact}");
                    }
                }
            }

            // === Phase 4：大運加成 ===
            if (!string.IsNullOrEmpty(daYunLiuShen))
            {
                string fullLiuShen = LiuShenFullMap.TryGetValue(daYunLiuShen, out var full) ? full : daYunLiuShen;
                var daYunRules = await _context.FortuneRules
                    .Where(r => r.Subcategory == "十神應事" && r.Title == fullLiuShen && r.IsActive)
                    .ToListAsync();

                if (daYunRules.Any())
                {
                    sb.AppendLine();
                    sb.AppendLine($"--- 大運加成（當前大運：{daYunStem}{daYunBranch} {fullLiuShen}）---");
                    var good = daYunRules.FirstOrDefault(r => r.ConditionText == "喜用")?.ResultText;
                    var bad  = daYunRules.FirstOrDefault(r => r.ConditionText == "忌用")?.ResultText;
                    if (good != null) sb.AppendLine($"【大運順勢】{good}");
                    if (bad  != null) sb.AppendLine($"【大運逆勢】{bad}");
                }
            }

            // === Phase 3：紫微命宮主星加成 ===

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

        /// <summary>依年齡取得今日運勢開頭提示（對應 ConsultationController.LfAgeTopicHint）</summary>
        private static string GetDailyAgeTopicHint(int age)
        {
            if (age <= 15) return "【年齡提示】學習成長期，今日重點：學業、家庭支持、品格養成。";
            if (age <= 25) return "【年齡提示】青年起步期，今日重點：事業初探、感情萌芽。";
            if (age <= 60) return "";
            if (age <= 70) return "【年齡提示】熟齡階段，今日重點：健康養護、財庫穩定、子女孫輩。";
            if (age <= 80) return "【年齡提示】長者階段，今日重點：健康長壽、財庫守護。";
            return "【年齡提示】高齡養生，今日重點：健康養生、財庫守護。";
        }

        /// <summary>依年齡取得 Gemini prompt 中的年齡注意事項</summary>
        private static string GetGeminiAgeNote(int age)
        {
            if (age <= 15) return "注意事項：命主年幼，禁止談感情婚姻，著重學業、家庭庇蔭、品格養成。";
            if (age >= 75) return "注意事項：命主高齡，禁止談感情婚姻，著重健康長壽、財庫守護。";
            if (age >= 61) return "注意事項：命主熟齡，著重健康財庫，感情部分可簡略。";
            return "";
        }

        private static readonly Dictionary<string, string> LiuShenFullMap = new()
        {
            {"比","比肩"},{"劫","劫財"},{"食","食神"},{"傷","傷官"},
            {"財","偏財"},{"才","正財"},{"殺","七殺"},{"官","正官"},
            {"梟","偏印"},{"印","正印"}
        };

        // 天干五行對應
        private static readonly Dictionary<string, string> StemElement = new()
        {
            {"甲","木"},{"乙","木"},{"丙","火"},{"丁","火"},
            {"戊","土"},{"己","土"},{"庚","金"},{"辛","金"},{"壬","水"},{"癸","水"}
        };

        // 五行相生：木→火→土→金→水→木
        private static readonly Dictionary<string, string> ElementGenerates = new()
            { {"木","火"},{"火","土"},{"土","金"},{"金","水"},{"水","木"} };

        // 地支藏干（主氣 8/5、中氣 3/2、餘氣 1，以權重表示）
        private static readonly Dictionary<string, Dictionary<string, int>> BranchHidden = new()
        {
            {"子", new(){{"癸",8}}},
            {"丑", new(){{"己",5},{"癸",2},{"辛",1}}},
            {"寅", new(){{"甲",5},{"丙",2},{"戊",1}}},
            {"卯", new(){{"乙",8}}},
            {"辰", new(){{"戊",5},{"乙",2},{"癸",1}}},
            {"巳", new(){{"丙",5},{"戊",2},{"庚",1}}},
            {"午", new(){{"丁",5},{"己",3}}},
            {"未", new(){{"己",5},{"丁",2},{"乙",1}}},
            {"申", new(){{"庚",5},{"壬",2},{"戊",1}}},
            {"酉", new(){{"辛",8}}},
            {"戌", new(){{"戊",5},{"辛",2},{"丁",1}}},
            {"亥", new(){{"壬",5},{"甲",3}}}
        };

        // 身強用神（順勢）：食傷財官殺
        private static readonly HashSet<string> StrongYongShen = new()
            { "食神","傷官","偏財","正財","七殺","正官" };

        // 身弱用神（順勢）：印比劫
        private static readonly HashSet<string> WeakYongShen = new()
            { "比肩","劫財","偏印","正印" };

        // 六沖組合（雙向）
        private static readonly HashSet<string> ChongPairs = new()
            { "子午","午子","丑未","未丑","寅申","申寅","卯酉","酉卯","辰戌","戌辰","巳亥","亥巳" };

        /// <summary>計算日主身強弱評分（正值越高越強）</summary>
        private static int CalcBodyStrength(string dayMaster, string monthBranch,
            IEnumerable<string> pillarStems, IEnumerable<string> pillarBranches)
        {
            int score = 0;
            if (!StemElement.TryGetValue(dayMaster, out var dmElem)) return 3;
            string genElem = ElementGenerates.FirstOrDefault(kv => kv.Value == dmElem).Key ?? "";

            // 月令（最大權重）
            if (BranchHidden.TryGetValue(monthBranch, out var monthH))
            {
                foreach (var (stem, w) in monthH)
                {
                    if (!StemElement.TryGetValue(stem, out var sElem)) continue;
                    if (sElem == dmElem)  score += w >= 5 ? 3 : w >= 2 ? 2 : 1; // 比劫
                    else if (sElem == genElem) score += w >= 5 ? 2 : 1;          // 印星
                }
            }

            // 年月時天干幫身
            foreach (var stem in pillarStems.Where(s => !string.IsNullOrEmpty(s)))
            {
                if (!StemElement.TryGetValue(stem, out var sElem)) continue;
                if (sElem == dmElem || sElem == genElem) score++;
            }

            // 地支通根
            var branches = pillarBranches.Where(b => !string.IsNullOrEmpty(b)).ToList();
            foreach (var br in branches)
            {
                if (!BranchHidden.TryGetValue(br, out var bh)) continue;
                if (bh.Any(kv => StemElement.TryGetValue(kv.Key, out var e) && (e == dmElem || e == genElem)))
                    score++;
            }

            // 命局內部六沖扣分（日支被沖-2，月支被沖-1）
            for (int i = 0; i < branches.Count; i++)
                for (int j = i + 1; j < branches.Count; j++)
                    if (ChongPairs.Contains(branches[i] + branches[j]))
                    {
                        if (i == 2 || j == 2) score -= 2;
                        else if (i == 1 || j == 1) score -= 1;
                    }

            // 三刑扣分
            var xingGroups = new[] { new[]{"寅","巳","申"}, new[]{"丑","戌","未"} };
            foreach (var g in xingGroups)
                if (g.Count(x => branches.Contains(x)) >= 2) score -= 1;

            return score;
        }

        /// <summary>從 baziInfo JSON 元素中取出單柱天干地支</summary>
        private static (string stem, string branch) ExtractPillar(JsonElement baziEl, string pillarKey)
        {
            JsonElement pillarEl = default;
            if (!baziEl.TryGetProperty(pillarKey, out pillarEl))
            {
                string cap = char.ToUpper(pillarKey[0]) + pillarKey[1..];
                if (!baziEl.TryGetProperty(cap, out pillarEl)) return ("", "");
            }
            string stem = "";
            string branch = "";
            if (pillarEl.TryGetProperty("heavenlyStem", out var hs) || pillarEl.TryGetProperty("HeavenlyStem", out hs))
                stem = hs.GetString() ?? "";
            if (pillarEl.TryGetProperty("earthlyBranch", out var eb) || pillarEl.TryGetProperty("EarthlyBranch", out eb))
                branch = eb.GetString() ?? "";
            return (stem, branch);
        }

        /// <summary>依關係類型與順逆勢取得影響描述</summary>
        private static string GetRelationImpact(string relationType, bool? isShun)
        {
            string dir = isShun == true ? "順" : isShun == false ? "逆" : "運";
            return (relationType, dir) switch
            {
                ("沖","順") => "能量受阻，好事多折，宜謹慎行事",
                ("沖","逆") => "壓力加劇，宜低調避衝",
                ("沖","運") => "氣場波動，宜穩不宜動",
                ("合","順") => "機遇受合化，方向可能轉變",
                ("合","逆") => "合化有緩解，壓力稍降",
                ("合","運") => "合化之日，注意方向轉變",
                ("刑","順") => "易生波折，計劃易被打斷",
                ("刑","逆") => "耗損加重，小心意外損傷",
                ("刑","運") => "暗藏波折，謹慎行事",
                ("破","順") => "成果易損，防虎頭蛇尾",
                ("破","逆") => "情況複雜，謹防小人",
                ("破","運") => "防事情破局",
                ("害","順") => "暗中阻滯，防無心之失",
                ("害","逆") => "小人干擾，謹言慎行",
                ("害","運") => "有隱性阻礙，留意人際",
                _ => "注意此日干支互動帶來的變數"
            };
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
