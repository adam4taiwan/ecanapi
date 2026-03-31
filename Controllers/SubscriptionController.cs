using Ecanapi.Data;
using Ecanapi.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Security.Claims;

namespace Ecanapi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class SubscriptionController : ControllerBase
    {
        private readonly ApplicationDbContext _context;
        private readonly IConfiguration _config;

        public SubscriptionController(ApplicationDbContext context, IConfiguration config)
        {
            _context = context;
            _config = config;
        }

        // ─── Public ──────────────────────────────────────────────────────────

        // GET /api/Subscription/plans
        // Returns all active membership plans with their benefits
        [HttpGet("plans")]
        public async Task<IActionResult> GetPlans()
        {
            var plans = await _context.MembershipPlans
                .Where(p => p.IsActive)
                .OrderBy(p => p.SortOrder)
                .Include(p => p.Benefits)
                .Select(p => new
                {
                    p.Id,
                    p.Code,
                    p.Name,
                    p.PriceTwd,
                    p.DurationDays,
                    p.Description,
                    Benefits = p.Benefits.Select(b => new
                    {
                        b.ProductCode,
                        b.ProductType,
                        b.BenefitType,
                        b.BenefitValue,
                        b.Description
                    })
                })
                .ToListAsync();

            return Ok(plans);
        }

        // GET /api/Subscription/products
        // Returns all active products (for display / pricing reference)
        [HttpGet("products")]
        public async Task<IActionResult> GetProducts()
        {
            var products = await _context.Products
                .Where(p => p.IsActive)
                .OrderBy(p => p.SortOrder)
                .Select(p => new
                {
                    p.Code,
                    p.Name,
                    p.Type,
                    p.PointCost,
                    p.PriceTwd,
                    p.Description
                })
                .ToListAsync();

            return Ok(products);
        }

        // ─── Authenticated ────────────────────────────────────────────────────

        // GET /api/Subscription/status
        // Returns the current user's active subscription and remaining quotas
        [Authorize]
        [HttpGet("status")]
        public async Task<IActionResult> GetStatus()
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (string.IsNullOrEmpty(userId)) return Unauthorized();

            // Admin 直接回傳最高權限，不受訂閱限制
            var user = await _context.Users.FindAsync(userId);
            var adminEmail = _config["Admin:Email"];
            if (user != null && string.Equals(user.Email, adminEmail, StringComparison.OrdinalIgnoreCase))
            {
                var adminProducts = new[] { "BOOK_BAZI", "BOOK_DAIYUN", "BOOK_LIUNIAN", "DAILY_FORTUNE", "BLESSING_ANTAISUI", "BLESSING_LIGHT", "BLESSING_WEALTH", "BLESSING_PRAYER", "CONSULT_VIDEO", "COURSE_BASIC", "LECTURE_FREE" };
                return Ok(new
                {
                    isSubscribed = true,
                    planCode = "GOLD",
                    planName = "系統管理員",
                    startDate = DateTime.UtcNow.Date,
                    expiryDate = DateTime.UtcNow.Date.AddYears(10),
                    daysRemaining = 3650,
                    birthdateLocked = false,
                    quotaStatus = adminProducts.Select(p => new
                    {
                        productCode = p,
                        productType = (string?)null,
                        total = 999,
                        used = 0,
                        remaining = 999
                    })
                });
            }

            var now = DateTime.UtcNow;
            var sub = await _context.UserSubscriptions
                .Where(s => s.UserId == userId && s.Status == "active" && s.ExpiryDate > now)
                .OrderByDescending(s => s.ExpiryDate)
                .Include(s => s.Plan)
                .ThenInclude(p => p.Benefits)
                .FirstOrDefaultAsync();

            // Check if birthdate is locked (any book report ever generated)
            var bookCodes = new[] { "BOOK_BAZI", "BOOK_DAIYUN", "BOOK_LIUNIAN" };
            var birthdateLocked = await _context.UserSubscriptionClaims
                .AnyAsync(c => c.UserId == userId && bookCodes.Contains(c.ProductCode));

            if (sub == null)
                return Ok(new { isSubscribed = false, birthdateLocked });

            // Calculate remaining quotas per product
            var quotaBenefits = sub.Plan.Benefits
                .Where(b => b.BenefitType == "quota")
                .ToList();

            var quotaStatus = new List<object>();
            foreach (var benefit in quotaBenefits)
            {
                var productCode = benefit.ProductCode ?? benefit.ProductType ?? "";
                int total = int.TryParse(benefit.BenefitValue, out var v) ? v : 0;

                // Annual quota: count claims in current calendar year
                var claimedThisYear = await _context.UserSubscriptionClaims
                    .CountAsync(c => c.UserId == userId
                                  && c.SubscriptionId == sub.Id
                                  && (c.ProductCode == productCode || (benefit.ProductType != null && c.ProductCode.StartsWith("BLESSING_")))
                                  && c.ClaimYear == now.Year);

                quotaStatus.Add(new
                {
                    ProductCode = productCode,
                    ProductType = benefit.ProductType,
                    Total = total,
                    Used = claimedThisYear,
                    Remaining = Math.Max(0, total - claimedThisYear)
                });
            }

            return Ok(new
            {
                isSubscribed = true,
                planCode = sub.Plan.Code,
                planName = sub.Plan.Name,
                startDate = sub.StartDate,
                expiryDate = sub.ExpiryDate,
                daysRemaining = (int)(sub.ExpiryDate - now).TotalDays,
                birthdateLocked,
                benefits = sub.Plan.Benefits.Select(b => new
                {
                    b.ProductCode,
                    b.ProductType,
                    b.BenefitType,
                    b.BenefitValue,
                    b.Description
                }),
                quotaStatus
            });
        }

        // POST /api/Subscription/claim
        // Claim a free quota item (e.g. annual liunian book or blessing)
        [Authorize]
        [HttpPost("claim")]
        public async Task<IActionResult> Claim([FromBody] ClaimRequest request)
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (string.IsNullOrEmpty(userId)) return Unauthorized();

            var now = DateTime.UtcNow;
            var sub = await _context.UserSubscriptions
                .Where(s => s.UserId == userId && s.Status == "active" && s.ExpiryDate > now)
                .OrderByDescending(s => s.ExpiryDate)
                .Include(s => s.Plan)
                .ThenInclude(p => p.Benefits)
                .FirstOrDefaultAsync();

            if (sub == null)
                return BadRequest(new { message = "尚無有效訂閱" });

            // Find a matching quota benefit for the requested product
            var product = await _context.Products.FirstOrDefaultAsync(p => p.Code == request.ProductCode && p.IsActive);
            if (product == null)
                return BadRequest(new { message = "商品不存在" });

            // Match by exact code or by type (for blessing wildcard)
            var benefit = sub.Plan.Benefits.FirstOrDefault(b =>
                b.BenefitType == "quota" &&
                (b.ProductCode == request.ProductCode || b.ProductType == product.Type));

            if (benefit == null)
                return BadRequest(new { message = "此訂閱方案不包含此商品的免費額度" });

            int quota = int.TryParse(benefit.BenefitValue, out var q) ? q : 0;

            // Count claims for this product type in the current year
            int usedThisYear = await _context.UserSubscriptionClaims
                .CountAsync(c => c.UserId == userId
                              && c.SubscriptionId == sub.Id
                              && (c.ProductCode == request.ProductCode ||
                                  (benefit.ProductType != null && c.ProductCode.StartsWith("BLESSING_") && product.Type == "blessing"))
                              && c.ClaimYear == now.Year);

            if (usedThisYear >= quota)
                return BadRequest(new { message = $"今年度 {product.Name} 免費額度已使用完畢" });

            _context.UserSubscriptionClaims.Add(new UserSubscriptionClaim
            {
                UserId         = userId,
                SubscriptionId = sub.Id,
                ProductCode    = request.ProductCode,
                ClaimYear      = now.Year,
                ClaimedAt      = now
            });

            await _context.SaveChangesAsync();
            return Ok(new { message = $"{product.Name} 免費額度已兌換" });
        }

        // ─── Dev-only: create a test subscription (admin only) ───────────────
        // POST /api/Subscription/dev-create?planCode=BRONZE|SILVER|GOLD
        [Authorize]
        [HttpPost("dev-create")]
        public async Task<IActionResult> DevCreateSubscription([FromQuery] string planCode = "SILVER")
        {
            // Only available in development environment
            var env = Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT") ?? "Production";
            if (env != "Development")
                return NotFound();

            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (string.IsNullOrEmpty(userId)) return Unauthorized();

            var plan = await _context.MembershipPlans.FirstOrDefaultAsync(p => p.Code == planCode && p.IsActive);
            if (plan == null) return BadRequest(new { message = $"Plan {planCode} not found" });

            var now = DateTime.UtcNow;
            // Remove existing active subscriptions for this user
            var existing = await _context.UserSubscriptions
                .Where(s => s.UserId == userId && s.Status == "active")
                .ToListAsync();
            foreach (var s in existing) s.Status = "cancelled";

            _context.UserSubscriptions.Add(new UserSubscription
            {
                UserId = userId,
                PlanId = plan.Id,
                StartDate = now,
                ExpiryDate = now.AddDays(365),
                Status = "active",
                PaymentRef = "DEV_TEST",
                CreatedAt = now
            });

            await _context.SaveChangesAsync();
            return Ok(new { message = $"Test subscription created: {plan.Name}", planCode = plan.Code });
        }

        // ─── Helper to get effective price for a user/product ────────────────

        // GET /api/Subscription/price?productCode=BOOK_BAZI
        // Returns the effective price (with any subscription discount applied)
        [Authorize]
        [HttpGet("price")]
        public async Task<IActionResult> GetEffectivePrice([FromQuery] string productCode)
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (string.IsNullOrEmpty(userId)) return Unauthorized();

            var product = await _context.Products
                .FirstOrDefaultAsync(p => p.Code == productCode && p.IsActive);
            if (product == null)
                return NotFound(new { message = "商品不存在" });

            var now = DateTime.UtcNow;
            var sub = await _context.UserSubscriptions
                .Where(s => s.UserId == userId && s.Status == "active" && s.ExpiryDate > now)
                .Include(s => s.Plan)
                .ThenInclude(p => p.Benefits)
                .FirstOrDefaultAsync();

            decimal discountRate = 1.0m;
            string? discountSource = null;

            if (sub != null)
            {
                // Check for exact product discount first, then type-level discount
                var exactDiscount = sub.Plan.Benefits.FirstOrDefault(b =>
                    b.BenefitType == "discount" && b.ProductCode == productCode);
                var typeDiscount = sub.Plan.Benefits.FirstOrDefault(b =>
                    b.BenefitType == "discount" && b.ProductType == product.Type);

                var best = exactDiscount ?? typeDiscount;
                if (best != null && decimal.TryParse(best.BenefitValue, out var rate))
                {
                    discountRate = rate;
                    discountSource = sub.Plan.Name;
                }
            }

            return Ok(new
            {
                productCode = product.Code,
                productName = product.Name,
                basePointCost = product.PointCost,
                basePriceTwd  = product.PriceTwd,
                discountRate,
                discountSource,
                effectivePointCost = product.PointCost.HasValue
                    ? (int)Math.Ceiling(product.PointCost.Value * discountRate) : (int?)null,
                effectivePriceTwd  = product.PriceTwd.HasValue
                    ? (int)Math.Ceiling(product.PriceTwd.Value * discountRate) : (int?)null
            });
        }
    }

    public class ClaimRequest
    {
        public string ProductCode { get; set; } = "";
    }
}
