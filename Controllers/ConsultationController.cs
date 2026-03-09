using Ecanapi.Data;
using Ecanapi.Models;
using Ecanapi.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Stripe;
using Stripe.Checkout;
using System.Security.Claims;
using System.Text.Json;

namespace Ecanapi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ConsultationController : ControllerBase
    {
        private readonly IAstrologyService _astrologyService;
        private readonly ApplicationDbContext _context;
        private readonly IConfiguration _config;
        private readonly ILogger<ConsultationController> _logger;
        private static readonly HttpClient _httpClient = new HttpClient { Timeout = TimeSpan.FromMinutes(5) };

        public ConsultationController(IAstrologyService astrologyService, ApplicationDbContext context, IConfiguration config, ILogger<ConsultationController> logger)
        {
            _astrologyService = astrologyService;
            _context = context;
            _config = config;
            _logger = logger;
        }

        [Authorize]
        [HttpPost("analyze")]
        public async Task<IActionResult> GetAiAnalysis([FromBody] AiRequest request)
        {
            var identity = User.FindFirstValue(ClaimTypes.Email)
                         ?? User.FindFirstValue(ClaimTypes.Name)
                         ?? User.FindFirst("unique_name")?.Value;

            if (string.IsNullOrEmpty(identity))
                return Unauthorized(new { error = "請重新登入" });

            var user = await _context.Users.FirstOrDefaultAsync(u => u.UserName == identity || u.Email == identity);
            if (user == null) return BadRequest(new { error = $"找不到用戶: {identity}" });

            if (request.Type == "同步查詢")
                return Ok(new { remainingPoints = user.Points });

            // 依類型與大運年數決定點數費用
            int cost = request.Type switch
            {
                "綜合性命書" or "綜合鑑定" => 50,
                "大運命書" => (request.FortuneDuration ?? 5) switch
                {
                    5 => 150,
                    10 => 200,
                    20 => 250,
                    30 => 300,
                    _ => 500   // 0 = 終身
                },
                "流年命書" => 20,
                "問事" => 10,
                _ => 50
            };

            if (user.Points < cost) return BadRequest(new { error = $"點數不足，此功能需要 {cost} 點" });

            try
            {
                var chartData = await _astrologyService.CalculateChartAsync(request.ChartRequest);
                string chartJson = JsonSerializer.Serialize(chartData);
                int nowYear = DateTime.Now.Year;
                int targetYear = request.TargetYear ?? nowYear;
                int birthYear = request.ChartRequest.Year;
                int currentAge = nowYear - birthYear;

                string prompt = request.Type switch
                {
                    "綜合性命書" or "綜合鑑定" => BuildComprehensivePrompt(chartJson),
                    "大運命書" => BuildMajorLuckPrompt(chartJson, chartData, currentAge, nowYear, request.FortuneDuration ?? 5),
                    "流年命書" => BuildAnnualLuckPrompt(chartJson, targetYear, targetYear - birthYear),
                    "問事" => BuildTopicPrompt(chartJson, request.Topic ?? ""),
                    _ => BuildComprehensivePrompt(chartJson)
                };

                string aiResult = await CallGeminiApi(prompt);

                user.Points -= cost;
                await _context.SaveChangesAsync();

                return Ok(new { result = aiResult, remainingPoints = user.Points });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "命理鑑定失敗 Type={Type} User={User}", request.Type, identity);
                return StatusCode(500, new { error = "命理鑑定失敗，請稍後再試", details = ex.Message });
            }
        }

        private string BuildComprehensivePrompt(string chartJson) => $@"
你是一位精通《三命通會》與《紫微斗數》的命理鑑定大師『玉洞子』。
請根據以下 JSON 命盤數據進行深度解析，並「嚴格」遵守下方的【模版排版規範】。

### 命盤數據：
{chartJson}

### 輸出格式：

一、格局判斷 (依訣評析)
【依據】
    * 八字解析：(描述日主生於何月、干支五行互動，需包含納音五行)
    * 紫微解析：(描述命、財、官、遷之星曜組合)
【格局】
    * 八字：『(格局名稱)』之真格。
    * 紫微：『(格局名稱)』之複合影響。

二、核心特質 (邏輯制化)
【定論】(用一句堅定的話斷死命格)
【解析】(深度解析內在與外在的互動)

三、關鍵斷語 (深度細查)
【財官運勢】：(明確斷言事業成就等級，嚴禁使用「可能」)
【六親情感】：(明確斷言婚姻配偶特徵)
【健康壽元】：(分析身體強弱、注意病符)

四、具體建議 (行為指南)
【行動指南】：(給出白話、具體、可執行的三條動作建議)

-----------------------------------------------------------------
鑑定大師：玉洞子  |  修身齊家，命在人心。
";

        private string BuildMajorLuckPrompt(string chartJson, AstrologyChartResult chartData, int currentAge, int nowYear, int duration)
        {
            // 計算年份範圍
            int endYear = duration == 0 ? nowYear + 80 : nowYear + duration;
            var annualLucks = _astrologyService.GenerateAnnualLucks(nowYear, endYear, chartData.Bazi.DayPillar.HeavenlyStem);
            string annualJson = JsonSerializer.Serialize(annualLucks);

            string durationLabel = duration == 0 ? "終身" : $"{duration}年";
            string detailLevel = duration switch
            {
                5 => "請逐月分析，表格欄位名稱固定為：月份、月柱、六神、吉凶、斷語。每格斷語內容不超過30字，表格標題列只寫欄位名稱。",
                10 => "請逐月分析，表格欄位名稱固定為：月份、月柱、六神、吉凶、斷語。每格斷語內容不超過20字，表格標題列只寫欄位名稱。",
                20 => "請逐季分析（每季3個月合併），表格欄位名稱固定為：季度、干支、六神、吉凶、重點。每格重點內容不超過30字，表格標題列只寫欄位名稱。",
                30 => "請逐年分析，表格欄位名稱固定為：年份、干支、六神、吉凶、重點。每格重點內容不超過30字，表格標題列只寫欄位名稱。",
                _ => "請逐大運分析，表格欄位名稱固定為：大運、干支、重點年份、吉凶、總評。每格總評不超過30字，表格標題列只寫欄位名稱。"
            };

            return $@"
你是一位精通《三命通會》大運推演與《紫微斗數》四化的命理鑑定大師『玉洞子』。
請根據以下命盤數據，從命主當前年齡（{currentAge} 歲，西元 {nowYear} 年）起，
出具【{durationLabel}大運命書】，依每個月吉凶分別判斷。

## 分析依據（務必逐一運用）：
1. 八字干支：合（六合、三合、半三合）、沖（六沖）、刑（三刑）、害（六害）、破（六破）
2. 神煞：天德、月德、天乙、文昌、文曲、羊刃、七煞、白虎、弔客、喪門等
3. 六神（十神）：每月月柱天干六神對日主的生剋制化
4. 紫微斗數：相關宮位四化（化祿/化權/化科/化忌）在大運/流年流月的觸發

## 命盤基礎數據：
{chartJson}

## 流年干支對照表（{nowYear}～{endYear}）：
{annualJson}

## 月柱推算規則（五虎遁年起月）：
甲己年→正月丙寅起；乙庚年→正月戊寅起；丙辛年→正月庚寅起；
丁壬年→正月壬寅起；戊癸年→正月甲寅起；月支：寅卯辰巳午未申酉戌亥子丑

## 輸出要求：
{detailLevel}

## 固定首尾格式：

大運命書 · {durationLabel} · 自 {nowYear} 年起

【大運總覽】
(先列出此期間經歷的大運干支、歲數區間、十神定性)

【逐月/逐季/逐年詳析】
(依上方「輸出要求」格式展開，吉月標▲、凶月標▼、平月標－)

【{durationLabel}行動總綱】
* 最佳出手年份：
* 最需低調年份：
* 財運重點：
* 婚姻重點：
* 事業重點：
* 健康提示：

-----------------------------------------------------------------
鑑定大師：玉洞子  |  知命善用，趨吉避凶。
";
        }

        private string BuildAnnualLuckPrompt(string chartJson, int targetYear, int currentAge) => $@"
你是一位精通流年推演的命理鑑定大師『玉洞子』。
請根據以下 JSON 命盤數據，針對 {targetYear} 年（命主約 {currentAge} 歲）進行深度流年分析。

### 命盤數據：
{chartJson}

### 輸出格式：

一、{targetYear} 年流年總論
【流年天干地支】：(說明該年干支)
【與命盤互動】：(流年干支對本命八字的生剋沖合)
【大運配合】：(此流年所在之大運，吉凶是否呼應)
【年運等級】：(★☆☆☆☆ 至 ★★★★★ + 一句斷語)

二、六大面向逐一研判
【財運】：(明確斷言財富進退，禁用「可能」)
【事業】：(明確斷言職場升遷或波折)
【感情婚姻】：(明確斷言感情動向)
【健康】：(點出需注意的身體部位或病符)
【人際是非】：(官司、小人、合作等)
【重大變動】：(搬遷、轉業、出國等可能性)

三、逐季重點提示
【第一季 (1-3月)】：
【第二季 (4-6月)】：
【第三季 (7-9月)】：
【第四季 (10-12月)】：

四、{targetYear} 年行動建議
【最佳時機】：(宜把握哪幾個月動作)
【謹慎時段】：(哪幾個月需低調)
【具體行動】：(三條可執行建議)

-----------------------------------------------------------------
鑑定大師：玉洞子  |  知運在先，趨吉避凶。
";

        private string BuildTopicPrompt(string chartJson, string topic) => $@"
你是一位精通《三命通會》與《紫微斗數》的命理鑑定大師『玉洞子』。
請根據以下 JSON 命盤數據，專門針對命主的【{topic}】問題進行深度鑑定，
給出明確、直接、可執行的命理建議。

### 命盤數據：
{chartJson}

### 問事主題：{topic}

### 輸出格式：

一、{topic} 命格基礎
【命盤顯示】：(此命盤在{topic}方面的先天條件，從八字與紫微雙角度分析)
【喜忌分析】：(哪些五行/星曜對{topic}有利，哪些有害)

二、{topic} 深度研判
【吉象】：(命盤中有利於{topic}的具體星曜/干支，明確說明)
【凶象】：(命盤中不利於{topic}的阻礙，明確說明)
【綜合評斷】：(用一句話斷死{topic}的整體格局，嚴禁使用「可能」或「或許」)

三、{topic} 時機研判
【最佳行動年份】：(從當前大運流年分析，哪幾年最適合)
【需要迴避時段】：(哪幾年不宜輕舉妄動)

四、{topic} 具體行動建議
【立即可做】：(現在就能執行的三件事)
【長期規劃】：(未來 3-5 年的策略方向)
【開運化煞】：(針對{topic}的具體風水或擇日建議)

-----------------------------------------------------------------
鑑定大師：玉洞子  |  問事明心，決策有據。
";

        [HttpPost("create-checkout-session")]
        public async Task<IActionResult> CreateCheckoutSession([FromBody] CreateCheckoutSessionRequest request)
        {
            StripeConfiguration.ApiKey = _config["Stripe:SecretKey"];
            var options = new SessionCreateOptions
            {
                PaymentMethodTypes = new List<string> { "card" },
                LineItems = new List<SessionLineItemOptions> {
                    new SessionLineItemOptions {
                        PriceData = new SessionLineItemPriceDataOptions {
                            UnitAmount = request.Amount,
                            Currency = "twd",
                            ProductData = new SessionLineItemPriceDataProductDataOptions { Name = $"{request.Points} 點數" },
                        },
                        Quantity = 1,
                    },
                },
                Mode = "payment",
                SuccessUrl = request.SuccessUrl + "?session_id={CHECKOUT_SESSION_ID}",
                CancelUrl = request.CancelUrl,
            };
            var service = new SessionService();
            var session = await service.CreateAsync(options);
            return Ok(new { id = session.Id, url = session.Url });
        }

        private async Task<string> CallGeminiApi(string prompt)
        {
            var apiKey = _config["Gemini:ApiKey"];
            var url = $"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={apiKey}";
            var payload = new { contents = new[] { new { parts = new[] { new { text = prompt } } } } };
            var response = await _httpClient.PostAsJsonAsync(url, payload);
            var rawJson = await response.Content.ReadAsStringAsync();
            _logger.LogInformation("Gemini HTTP {StatusCode}: {Preview}", (int)response.StatusCode, rawJson.Length > 300 ? rawJson[..300] : rawJson);
            var json = JsonSerializer.Deserialize<JsonElement>(rawJson);
            if (!json.TryGetProperty("candidates", out var candidates))
                throw new Exception($"Gemini 回傳錯誤: {rawJson[..Math.Min(300, rawJson.Length)]}");
            return candidates[0].GetProperty("content").GetProperty("parts")[0].GetProperty("text").GetString();
        }
    }

    public class AiRequest
    {
        public string Type { get; set; }
        public AstrologyRequest ChartRequest { get; set; }
        public int? TargetYear { get; set; }        // 流年命書用
        public string? Topic { get; set; }           // 問事用
        public int? FortuneDuration { get; set; }    // 大運命書年數：5/10/20/30/0=終身
    }

    public class CreateCheckoutSessionRequest
    {
        public string UserId { get; set; }
        public int Amount { get; set; }
        public int Points { get; set; }
        public string SuccessUrl { get; set; }
        public string CancelUrl { get; set; }
    }

    public class PointRecord
    {
        public int Id { get; set; }
        public string UserId { get; set; }
        public int Amount { get; set; }
        public string? Description { get; set; }
        public string? StripeSessionId { get; set; }
        public DateTime CreatedAt { get; set; }
    }
}