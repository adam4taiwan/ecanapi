using Ecanapi.Data;
using Ecanapi.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Stripe;
using Stripe.Checkout;

namespace Ecanapi.Controllers
{
    // 確保 UserName 在 Swagger 可以輸入
    public class PointPackage
    {
        public int Points { get; set; }
        public int Price { get; set; }
        public string UserName { get; set; }
    }

    [ApiController]
    [Route("api/[controller]")]
    public class PaymentController : ControllerBase
    {
        private readonly ApplicationDbContext _context;
        private readonly IConfiguration _config;

        public PaymentController(ApplicationDbContext context, IConfiguration config)
        {
            _context = context;
            _config = config;
        }

        [HttpPost("create-checkout-session")]
        public async Task<IActionResult> CreateCheckoutSession([FromBody] PointPackage package)
        {
            StripeConfiguration.ApiKey = _config["Stripe:SecretKey"];
            // 🚩 這裡動態判斷：如果是本地就回 3000，如果是雲端就回您的 Fly.io 網址
            var origin = Request.Headers["Origin"].ToString();
            // 如果前端沒傳 Origin，就預設一個回傳路徑
            var frontendUrl = string.IsNullOrEmpty(origin) ? $"{HttpContext.Request.Scheme}://{HttpContext.Request.Host}" : origin;

            // 確保回到命盤頁面
            //var successUrl = $"{frontendUrl}/disk";
            // 🚩 修改目的：在網址後方加上一個隨機的時間戳記 (Timestamp)
            // 這樣瀏覽器會認為這是一個「全新的頁面」，從而強制前端執行 useEffect 抓取最新點數
            var successUrl = $"{frontendUrl}/disk?session_id={{CHECKOUT_SESSION_ID}}&t={DateTime.Now.Ticks}";
            // 自動抓取當前 Host，避免 Port 變動導致 404
            var host = HttpContext.Request.Host;
            var scheme = HttpContext.Request.Scheme;

            var options = new SessionCreateOptions
            {
                PaymentMethodTypes = new List<string> { "card" },
                LineItems = new List<SessionLineItemOptions> {
                    new SessionLineItemOptions {
                        PriceData = new SessionLineItemPriceDataOptions {
                            Currency = "twd",
                            UnitAmount = package.Price * 100,
                            ProductData = new SessionLineItemPriceDataProductDataOptions { Name = "點數充值" },
                        },
                        Quantity = 1,
                    },
                },
                Mode = "payment",
                SuccessUrl = successUrl,
                CancelUrl = successUrl,
                Metadata = new Dictionary<string, string> {
                    { "UserName", package.UserName },
                    { "Points", package.Points.ToString() }
                }
            };

            var service = new SessionService();
            var session = await service.CreateAsync(options);
            return Ok(new { url = session.Url });
        }

        [HttpPost("webhook")]
        public async Task<IActionResult> StripeWebhook()
        {
            var json = await new StreamReader(HttpContext.Request.Body).ReadToEndAsync();
            // 在這裡點紅點 (斷點)
            Console.WriteLine(">>>> [Webhook] 收到請求");

            try
            {
                var stripeEvent = EventUtility.ConstructEvent(
                    json, Request.Headers["Stripe-Signature"], _config["Stripe:WebhookSecret"]);

                if (stripeEvent.Type == "checkout.session.completed")
                {
                    var session = stripeEvent.Data.Object as Stripe.Checkout.Session;
                    var loginName = session.Metadata["UserName"];
                    var pointsToAdd = int.Parse(session.Metadata["Points"]);

                    Console.WriteLine($">>>> [Webhook] 準備為用戶 {loginName} 增加 {pointsToAdd} 點");

                    // 注意：這裡的大小寫必須與資料庫一致
                    // 🚩 在 PaymentController.cs 的 Webhook 處修改這行：
                    // 原本是：u.UserName == loginName
                    // 修改為：
                    var user = await _context.Users.FirstOrDefaultAsync(u => u.Email == loginName || u.UserName == loginName);
                    //var user = await _context.Users.FirstOrDefaultAsync(u => u.UserName == loginName);
                    if (user != null)
                    {
                        user.Points += pointsToAdd;
                        await _context.SaveChangesAsync();
                        Console.WriteLine(">>>> [Webhook] 資料庫更新成功！");
                    }
                    else
                    {
                        Console.WriteLine($">>>> [Webhook] 錯誤：找不到用戶 {loginName}");
                    }
                }
                return Ok();
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($">>>> [Webhook] 驗證失敗: {ex.Message}");
                return BadRequest();
            }
        }
    }
}