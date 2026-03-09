using Ecanapi.Data;
using Ecanapi.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Stripe;
using Stripe.Checkout;
using System.Security.Claims;

namespace Ecanapi.Controllers
{
    public class CheckoutRequest
    {
        public string PackageId { get; set; } = string.Empty;
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
