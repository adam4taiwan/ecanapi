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
        private static readonly HttpClient _httpClient = new HttpClient();

        public ConsultationController(IAstrologyService astrologyService, ApplicationDbContext context, IConfiguration config)
        {
            _astrologyService = astrologyService;
            _context = context;
            _config = config;
        }

        [Authorize]
        [HttpPost("analyze")]
        public async Task<IActionResult> GetAiAnalysis([FromBody] AiRequest request)
        {
            // 🚩 解決 UserName 字典報錯：改用三合一讀取身分
            var identity = User.FindFirstValue(ClaimTypes.Email)
                         ?? User.FindFirstValue(ClaimTypes.Name)
                         ?? User.FindFirst("unique_name")?.Value;

            if (string.IsNullOrEmpty(identity))
                return Unauthorized(new { error = "請重新登入" });

            // 同時比對 UserName 與 Email 欄位，確保 user 不會是 null
            var user = await _context.Users.FirstOrDefaultAsync(u => u.UserName == identity || u.Email == identity);
            if (user == null) return BadRequest(new { error = $"找不到用戶: {identity}" });

            // 同步查詢逻辑：如果 Type 是同步查詢，直接回傳點數，不進 AI
            if (request.Type == "同步查詢")
            {
                return Ok(new { remainingPoints = user.Points });
            }

            if (user.Points < 50) return BadRequest(new { error = "點數不足" });

            // 檢查點數
            int cost = 50;

            try
            {
                // 對齊 Service 裡的 AstrologyRequest
                var chartData = await _astrologyService.CalculateChartAsync(request.ChartRequest);
                string chartJson = JsonSerializer.Serialize(chartData);

                // 🚩 核心優化：完全依照「模版格式.docx」設計的 Prompt
                string prompt = $@"
你是一位精通《三命通會》與《紫微斗數》的命理鑑定大師『玉洞子』。
請根據以下 JSON 命盤數據進行深度解析，並「嚴格」遵守下方的【模版排版規範】。

### 命盤數據：
{chartJson}

### 模版排版規範 (請直接模仿此格式輸出)：


一、格局判斷 (依訣評析)
【依據】 
    * 八字解析：(請模仿陳敢維範例，描述日主生於何月、干支五行互動，需包含納音五行)
    * 紫微解析：(描述命、財、官、遷之星曜組合，如：天相坐未、紫破對照)
【格局】
    * 八字：『(格局名稱)』之真格。
    * 紫微：『(格局名稱)』之複合影響。

二、核心特質 (邏輯制化) 
【定論】 
    (用一句堅定的話斷死命格，如：性格剛毅但流於固執，凡事需親力親為。)
【解析】
    (深度解析內在與外在的互動，如：外表敦厚而內在強悍，具備開創潛力。)

三、關鍵斷語 (深度細查)
【財官運勢】： (明確斷言事業成就等級，嚴禁使用「可能」)
【六親情感】： (明確斷言婚姻配偶特徵，如：娶妻強人、防產厄)

四、具體建議 (行為指南)
【行動指南】： (給出白話、具體、可執行的三條動作建議)

-----------------------------------------------------------------
鑑定大師：玉洞子  |  修身齊家，命在人心。
";

                string aiResult = await CallGeminiApi(prompt);

                // 扣除點數
                user.Points -= cost;
                await _context.SaveChangesAsync();

                // 🚩 處理排版：確保換行符號在前端顯示
                return Ok(new
                {
                    result = aiResult,
                    remainingPoints = user.Points
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = "AI 分析失敗", details = ex.Message });
            }
        }

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
            var json = await response.Content.ReadFromJsonAsync<JsonElement>();
            return json.GetProperty("candidates")[0].GetProperty("content").GetProperty("parts")[0].GetProperty("text").GetString();
        }
    }

    public class AiRequest
    {
        public string Type { get; set; }
        public AstrologyRequest ChartRequest { get; set; }
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
        public DateTime CreatedAt { get; set; }
    }
}