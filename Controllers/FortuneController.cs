using Ecanapi.Data;
using Ecanapi.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Security.Claims;
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
