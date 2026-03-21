using Ecanapi.Data;
using Ecanapi.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Stripe;
using Stripe.Checkout;
using System.Security.Claims;
using System.Security.Cryptography;
using System.Text;
using System.Web;

namespace Ecanapi.Controllers
{
    public class CheckoutRequest
    {
        public string PackageId { get; set; } = string.Empty;
    }

    public class SubscriptionCheckoutRequest
    {
        public string PlanCode { get; set; } = string.Empty;
    }

    public class AtmRequest
    {
        public string PackageId { get; set; } = string.Empty;
        public string TransferDate { get; set; } = string.Empty;
        public string AccountLast5 { get; set; } = string.Empty;
    }

    [ApiController]
    [Route("api/[controller]")]
    public class PaymentController : ControllerBase
    {
        private readonly ApplicationDbContext _context;
        private readonly IConfiguration _config;

        // Server-side defined valid packages — clients only pass packageId
        private static readonly Dictionary<string, (int Points, int PriceTwd)> ValidPackages = new()
        {
            { "starter",  (50,   500)  },
            { "popular",  (150,  1350) },
            { "advanced", (400,  3200) },
            { "vip",      (1000, 7000) },
        };

        public PaymentController(ApplicationDbContext context, IConfiguration config)
        {
            _context = context;
            _config = config;
        }

        [HttpGet("packages")]
        public IActionResult GetPackages()
        {
            var packages = ValidPackages.Select(p => new
            {
                id = p.Key,
                points = p.Value.Points,
                priceTwd = p.Value.PriceTwd,
            });
            return Ok(packages);
        }

        [Authorize]
        [HttpPost("create-checkout-session")]
        public async Task<IActionResult> CreateCheckoutSession([FromBody] CheckoutRequest request)
        {
            if (!ValidPackages.TryGetValue(request.PackageId, out var pkg))
                return BadRequest(new { message = "無效的點數套餐" });

            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (string.IsNullOrEmpty(userId))
                return Unauthorized();

            var user = await _context.Users.FindAsync(userId);
            if (user == null)
                return Unauthorized();

            StripeConfiguration.ApiKey = _config["Stripe:SecretKey"];

            var origin = Request.Headers["Origin"].ToString();
            var frontendUrl = string.IsNullOrEmpty(origin)
                ? $"{HttpContext.Request.Scheme}://{HttpContext.Request.Host}"
                : origin;

            var successUrl = $"{frontendUrl}/member?payment=success&session_id={{CHECKOUT_SESSION_ID}}";
            var cancelUrl  = $"{frontendUrl}/member?payment=cancelled";

            var options = new SessionCreateOptions
            {
                PaymentMethodTypes = new List<string> { "card" },
                LineItems = new List<SessionLineItemOptions>
                {
                    new SessionLineItemOptions
                    {
                        PriceData = new SessionLineItemPriceDataOptions
                        {
                            Currency = "twd",
                            UnitAmount = pkg.PriceTwd * 100,
                            ProductData = new SessionLineItemPriceDataProductDataOptions
                            {
                                Name = $"點數儲值 {pkg.Points} 點",
                            },
                        },
                        Quantity = 1,
                    },
                },
                Mode = "payment",
                SuccessUrl = successUrl,
                CancelUrl  = cancelUrl,
                Metadata = new Dictionary<string, string>
                {
                    { "UserId",    userId },
                    { "Points",    pkg.Points.ToString() },
                    { "PackageId", request.PackageId },
                },
            };

            var service = new SessionService();
            var session = await service.CreateAsync(options);
            return Ok(new { url = session.Url });
        }

        [Authorize]
        [HttpPost("atm-request")]
        public async Task<IActionResult> CreateAtmRequest([FromBody] AtmRequest request)
        {
            if (!ValidPackages.TryGetValue(request.PackageId, out var pkg))
                return BadRequest(new { message = "無效的點數套餐" });
            if (string.IsNullOrWhiteSpace(request.TransferDate) || string.IsNullOrWhiteSpace(request.AccountLast5))
                return BadRequest(new { message = "請填寫轉帳日期及帳號後 5 碼" });

            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (string.IsNullOrEmpty(userId)) return Unauthorized();

            _context.AtmPaymentRequests.Add(new AtmPaymentRequest
            {
                UserId      = userId,
                PackageId   = request.PackageId,
                Points      = pkg.Points,
                PriceTwd    = pkg.PriceTwd,
                TransferDate  = request.TransferDate,
                AccountLast5  = request.AccountLast5,
                Status      = "pending",
                CreatedAt   = DateTime.UtcNow,
            });
            await _context.SaveChangesAsync();
            return Ok(new { message = "ATM 轉帳申請已提交，審核後點數將自動入帳" });
        }

        // ─── ECPay ────────────────────────────────────────────────────────────

        private string EcpayCheckMac(Dictionary<string, string> parameters)
        {
            var hashKey = _config["ECPay:HashKey"] ?? "5294y06JbISpM5x9";
            var hashIV  = _config["ECPay:HashIV"]  ?? "v77hoKGq4kWxNNIS";

            var sorted = parameters
                .Where(p => p.Key != "CheckMacValue")
                .OrderBy(p => p.Key, StringComparer.OrdinalIgnoreCase);

            var raw = "HashKey=" + hashKey + "&" +
                      string.Join("&", sorted.Select(p => $"{p.Key}={p.Value}")) +
                      "&HashIV=" + hashIV;

            var encoded = HttpUtility.UrlEncode(raw).ToLower();

            var hash = SHA256.HashData(Encoding.UTF8.GetBytes(encoded));
            return Convert.ToHexString(hash).ToUpper();
        }

        [Authorize]
        [HttpPost("create-ecpay-checkout")]
        public async Task<IActionResult> CreateEcpayCheckout([FromBody] CheckoutRequest request)
        {
            if (!(_config.GetValue<bool?>("ECPay:Enabled") ?? false))
                return BadRequest(new { message = "儲值功能暫時停用，請稍後再試" });

            if (!ValidPackages.TryGetValue(request.PackageId, out var pkg))
                return BadRequest(new { message = "無效的點數套餐" });

            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (string.IsNullOrEmpty(userId)) return Unauthorized();

            var user = await _context.Users.FindAsync(userId);
            if (user == null) return Unauthorized();

            var tradeNo  = $"EC{DateTime.Now:yyyyMMddHHmmss}{new Random().Next(10, 99)}";
            var tradeDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

            var origin = Request.Headers["Origin"].ToString();
            var frontendUrl = string.IsNullOrEmpty(origin)
                ? $"{HttpContext.Request.Scheme}://{HttpContext.Request.Host}"
                : origin;

            var returnUrl      = _config["ECPay:ReturnUrl"] ?? "https://ecanapi.fly.dev/api/Payment/ecpay-return";
            var ecpapiBase     = _config["ECPay:ApiBase"] ?? "https://ecanapi.fly.dev";
            var orderResultUrl = $"{ecpapiBase}/api/Payment/ecpay-result?redirect={Uri.EscapeDataString($"{frontendUrl}/disk")}";

            var parameters = new Dictionary<string, string>
            {
                { "MerchantID",        _config["ECPay:MerchantID"] ?? "2000132" },
                { "MerchantTradeNo",   tradeNo },
                { "MerchantTradeDate", tradeDate },
                { "PaymentType",       "aio" },
                { "TotalAmount",       pkg.PriceTwd.ToString() },
                { "TradeDesc",         "Astrology Points" },
                { "ItemName",          $"Points {pkg.Points}" },
                { "ReturnURL",         returnUrl },
                { "OrderResultURL",    orderResultUrl },
                { "ChoosePayment",     "Credit" },
                { "EncryptType",       "1" },
                { "CustomField1",      userId },
                { "CustomField2",      pkg.Points.ToString() },
                { "CustomField3",      request.PackageId },
            };

            parameters["CheckMacValue"] = EcpayCheckMac(parameters);

            var paymentUrl = _config["ECPay:PaymentUrl"]
                ?? "https://payment-stage.ecpay.com.tw/Cashier/AioCheckout/V5";

            return Ok(new { actionUrl = paymentUrl, parameters });
        }

        // ─── ECPay Subscription Checkout ──────────────────────────────────────

        [Authorize]
        [HttpPost("create-subscription-checkout")]
        public async Task<IActionResult> CreateSubscriptionCheckout([FromBody] SubscriptionCheckoutRequest request)
        {
            if (!(_config.GetValue<bool?>("ECPay:Enabled") ?? false))
                return BadRequest(new { message = "訂閱付款功能即將開放，敬請期待" });

            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (string.IsNullOrEmpty(userId)) return Unauthorized();

            var plan = await _context.MembershipPlans
                .FirstOrDefaultAsync(p => p.Code == request.PlanCode && p.IsActive);
            if (plan == null) return BadRequest(new { message = "訂閱方案不存在" });

            var tradeNo   = $"SUB{DateTime.Now:yyyyMMddHHmmss}{new Random().Next(10, 99)}";
            var tradeDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

            var origin = Request.Headers["Origin"].ToString();
            var frontendUrl = string.IsNullOrEmpty(origin)
                ? $"{HttpContext.Request.Scheme}://{HttpContext.Request.Host}"
                : origin;

            var returnUrl      = _config["ECPay:ReturnUrl"] ?? "https://ecanapi.fly.dev/api/Payment/ecpay-return";
            var ecpapiBase     = _config["ECPay:ApiBase"] ?? "https://ecanapi.fly.dev";
            var orderResultUrl = $"{ecpapiBase}/api/Payment/ecpay-result?redirect={Uri.EscapeDataString($"{frontendUrl}/member")}";

            var parameters = new Dictionary<string, string>
            {
                { "MerchantID",        _config["ECPay:MerchantID"] ?? "2000132" },
                { "MerchantTradeNo",   tradeNo },
                { "MerchantTradeDate", tradeDate },
                { "PaymentType",       "aio" },
                { "TotalAmount",       plan.PriceTwd.ToString() },
                { "TradeDesc",         "Membership Subscription" },
                { "ItemName",          $"{plan.Name} 年訂閱" },
                { "ReturnURL",         returnUrl },
                { "OrderResultURL",    orderResultUrl },
                { "ChoosePayment",     "Credit" },
                { "EncryptType",       "1" },
                { "CustomField1",      userId },
                { "CustomField2",      plan.DurationDays.ToString() },
                { "CustomField3",      $"SUB_{plan.Code}" },
            };

            parameters["CheckMacValue"] = EcpayCheckMac(parameters);

            var paymentUrl = _config["ECPay:PaymentUrl"]
                ?? "https://payment-stage.ecpay.com.tw/Cashier/AioCheckout/V5";

            return Ok(new { actionUrl = paymentUrl, parameters });
        }

        [HttpPost("ecpay-return")]
        public async Task<IActionResult> EcpayReturn([FromForm] IFormCollection form)
        {
            if (!(_config.GetValue<bool?>("ECPay:Enabled") ?? false))
                return Content("1|OK");

            var formParams = form.ToDictionary(k => k.Key, v => v.Value.ToString());

            var receivedMac = formParams.GetValueOrDefault("CheckMacValue", "");
            formParams.Remove("CheckMacValue");
            var expectedMac = EcpayCheckMac(formParams);

            if (!string.Equals(receivedMac, expectedMac, StringComparison.OrdinalIgnoreCase))
                return Content("0|CheckMacValue Error");

            var rtnCode = formParams.GetValueOrDefault("RtnCode", "0");
            if (rtnCode != "1") return Content("1|OK");

            var tradeNo   = formParams.GetValueOrDefault("MerchantTradeNo", "");
            var userId    = formParams.GetValueOrDefault("CustomField1", "");
            var field2    = formParams.GetValueOrDefault("CustomField2", "");
            var field3    = formParams.GetValueOrDefault("CustomField3", "");

            var user = await _context.Users.FindAsync(userId);
            if (user == null) return Content("0|User Not Found");

            // ── Subscription purchase (CustomField3 starts with "SUB_") ───────
            if (field3.StartsWith("SUB_"))
            {
                var planCode = field3.Substring(4); // e.g. "GOLD"
                var alreadyProcessed = await _context.UserSubscriptions
                    .AnyAsync(s => s.PaymentRef == tradeNo);
                if (alreadyProcessed) return Content("1|OK");

                var plan = await _context.MembershipPlans
                    .FirstOrDefaultAsync(p => p.Code == planCode && p.IsActive);
                if (plan == null) return Content("0|Plan Not Found");

                var now = DateTime.UtcNow;
                // Extend from existing expiry if still active, else from now
                var existingSub = await _context.UserSubscriptions
                    .Where(s => s.UserId == userId && s.Status == "active" && s.ExpiryDate > now)
                    .OrderByDescending(s => s.ExpiryDate)
                    .FirstOrDefaultAsync();
                var startDate = existingSub?.ExpiryDate ?? now;

                _context.UserSubscriptions.Add(new UserSubscription
                {
                    UserId     = userId,
                    PlanId     = plan.Id,
                    StartDate  = startDate,
                    ExpiryDate = startDate.AddDays(plan.DurationDays),
                    Status     = "active",
                    PaymentRef = tradeNo,
                    CreatedAt  = now
                });
                await _context.SaveChangesAsync();
                return Content("1|OK");
            }

            // ── Point purchase ────────────────────────────────────────────────
            var packageId = field3;
            if (!int.TryParse(field2, out var pointsToAdd) || pointsToAdd <= 0)
                return Content("0|Invalid Points");

            var alreadyPointProcessed = await _context.PointRecords.AnyAsync(r => r.StripeSessionId == tradeNo);
            if (alreadyPointProcessed) return Content("1|OK");

            user.Points += pointsToAdd;
            _context.PointRecords.Add(new PointRecord
            {
                UserId          = userId,
                Amount          = pointsToAdd,
                Description     = $"ECPay {packageId} 套餐儲值",
                StripeSessionId = tradeNo,
                CreatedAt       = DateTime.UtcNow,
            });

            await _context.SaveChangesAsync();
            return Content("1|OK");
        }

        [HttpGet("ecpay-result")]
        [HttpPost("ecpay-result")]
        public IActionResult EcpayResult([FromQuery] string redirect = "")
        {
            var target = string.IsNullOrEmpty(redirect) ? "/" : $"{redirect}?payment=success";
            // Return HTML page that auto-redirects (works for both GET and POST)
            var html = $@"<!DOCTYPE html><html><head>
<meta http-equiv=""refresh"" content=""0;url={target}"">
<script>window.location.href='{target}';</script>
</head><body>付款完成，正在跳轉...</body></html>";
            return Content(html, "text/html");
        }

        // ─── Stripe (保留不動) ───────────────────────────────────────────────

        [HttpPost("webhook")]
        public async Task<IActionResult> StripeWebhook()
        {
            var json = await new StreamReader(HttpContext.Request.Body).ReadToEndAsync();
            Console.WriteLine(">>>> [Webhook] Received request");

            try
            {
                var stripeEvent = EventUtility.ConstructEvent(
                    json,
                    Request.Headers["Stripe-Signature"],
                    _config["Stripe:WebhookSecret"]);

                if (stripeEvent.Type == "checkout.session.completed")
                {
                    var session = stripeEvent.Data.Object as Stripe.Checkout.Session;
                    if (session == null) return Ok();

                    // Idempotency: skip if already processed
                    var alreadyProcessed = await _context.PointRecords
                        .AnyAsync(r => r.StripeSessionId == session.Id);
                    if (alreadyProcessed)
                    {
                        Console.WriteLine($">>>> [Webhook] Session {session.Id} already processed, skipping");
                        return Ok();
                    }

                    var userId    = session.Metadata["UserId"];
                    var pointsToAdd = int.Parse(session.Metadata["Points"]);
                    var packageId = session.Metadata.GetValueOrDefault("PackageId", "");

                    Console.WriteLine($">>>> [Webhook] Adding {pointsToAdd} points to userId={userId}");

                    var user = await _context.Users.FindAsync(userId);
                    if (user == null)
                    {
                        Console.WriteLine($">>>> [Webhook] Error: user not found userId={userId}");
                        return Ok();
                    }

                    user.Points += pointsToAdd;
                    _context.PointRecords.Add(new PointRecord
                    {
                        UserId          = userId,
                        Amount          = pointsToAdd,
                        Description     = $"Stripe {packageId} 套餐儲值",
                        StripeSessionId = session.Id,
                        CreatedAt       = DateTime.UtcNow,
                    });

                    await _context.SaveChangesAsync();
                    Console.WriteLine($">>>> [Webhook] Success: userId={userId} points={user.Points}");
                }

                return Ok();
            }
            catch (Exception ex)
            {
                Console.WriteLine($">>>> [Webhook] Signature verification failed: {ex.Message}");
                return BadRequest();
            }
        }
    }
}
