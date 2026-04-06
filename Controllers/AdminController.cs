using Ecanapi.Data;
using Ecanapi.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Security.Claims;

namespace Ecanapi.Controllers
{
    public class AdminUserDto
    {
        public string Id { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
        public string? Phone { get; set; }
        public string? PostalCode { get; set; }
        public string? Address { get; set; }
        public string? TaxId { get; set; }
        public int Points { get; set; }
        public int? BirthYear { get; set; }
        public int? BirthMonth { get; set; }
        public int? BirthDay { get; set; }
        public int? BirthHour { get; set; }
        public int? BirthGender { get; set; }
        public AdminSubscriptionDto? Subscription { get; set; }
    }

    public class AdminSubscriptionDto
    {
        public string PlanCode { get; set; } = string.Empty;
        public string PlanName { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public DateTime? ExpiryDate { get; set; }
        public bool IsInTrial { get; set; }
        public DateTime? TrialStartDate { get; set; }
        public int? TrialDaysRemaining { get; set; }
    }

    public class UpdateUserRequest
    {
        public string? Name { get; set; }
        public string? Phone { get; set; }
        public string? PostalCode { get; set; }
        public string? Address { get; set; }
        public string? TaxId { get; set; }
    }

    public class ForceChangePasswordRequest
    {
        public string NewPassword { get; set; } = string.Empty;
    }

    public class AdjustPointsRequest
    {
        public int Amount { get; set; }
        public string Reason { get; set; } = string.Empty;
    }

    public class AtmApproveRequest
    {
        public string? Note { get; set; }
    }

    public class AtmRejectRequest
    {
        public string Note { get; set; } = string.Empty;
    }

    [ApiController]
    [Route("api/[controller]")]
    [Authorize]
    public class AdminController : ControllerBase
    {
        private readonly ApplicationDbContext _context;
        private readonly UserManager<ApplicationUser> _userManager;
        private readonly IConfiguration _config;

        public AdminController(
            ApplicationDbContext context,
            UserManager<ApplicationUser> userManager,
            IConfiguration config)
        {
            _context = context;
            _userManager = userManager;
            _config = config;
        }

        private bool IsAdmin()
        {
            var email = User.FindFirstValue(ClaimTypes.Email);
            var adminEmail = _config["Admin:Email"];
            return !string.IsNullOrEmpty(email) && email == adminEmail;
        }

        // GET /api/Admin/verify
        [HttpGet("verify")]
        public IActionResult Verify()
        {
            if (!IsAdmin()) return Forbid();
            return Ok(new { isAdmin = true });
        }

        // DELETE /api/Admin/fortune-cache-today - 清除今日所有運勢快取（測試用）
        [HttpDelete("fortune-cache-today")]
        public async Task<IActionResult> ClearFortuneCacheToday()
        {
            if (!IsAdmin()) return Forbid();
            var today = DateTime.UtcNow.Date;
            var records = _context.DailyFortunes.Where(f => f.FortuneDate == today);
            var count = await records.CountAsync();
            _context.DailyFortunes.RemoveRange(records);
            await _context.SaveChangesAsync();
            return Ok(new { deleted = count, date = today.ToString("yyyy-MM-dd") });
        }

        // GET /api/Admin/stats
        [HttpGet("stats")]
        public async Task<IActionResult> GetStats()
        {
            if (!IsAdmin()) return Forbid();
            var totalUsers = await _userManager.Users.CountAsync();
            var totalPointsSold = await _context.PointRecords
                .Where(r => r.Amount > 0)
                .SumAsync(r => (int?)r.Amount) ?? 0;
            var pendingAtm = await _context.AtmPaymentRequests
                .CountAsync(r => r.Status == "pending");
            return Ok(new { totalUsers, totalPointsSold, pendingAtm });
        }

        // GET /api/Admin/users?search=
        [HttpGet("users")]
        public async Task<IActionResult> GetUsers([FromQuery] string? search)
        {
            if (!IsAdmin()) return Forbid();
            var query = _userManager.Users.AsQueryable();
            if (!string.IsNullOrEmpty(search))
                query = query.Where(u =>
                    (u.Email != null && u.Email.Contains(search)) ||
                    u.Name.Contains(search));
            var users = await query
                .OrderBy(u => u.Email)
                .Select(u => new AdminUserDto
                {
                    Id = u.Id,
                    Name = u.Name,
                    Email = u.Email ?? string.Empty,
                    Phone = u.Phone,
                    PostalCode = u.PostalCode,
                    Address = u.Address,
                    TaxId = u.TaxId,
                    Points = u.Points,
                    BirthYear = u.BirthYear,
                    BirthMonth = u.BirthMonth,
                    BirthDay = u.BirthDay,
                    BirthHour = u.BirthHour,
                    BirthGender = u.BirthGender,
                })
                .ToListAsync();
            return Ok(users);
        }

        // GET /api/Admin/users/{id}
        [HttpGet("users/{id}")]
        public async Task<IActionResult> GetUser(string id)
        {
            if (!IsAdmin()) return Forbid();
            var u = await _userManager.FindByIdAsync(id);
            if (u == null) return NotFound();

            var now = DateTime.UtcNow;
            var sub = await _context.UserSubscriptions
                .Include(s => s.Plan)
                .Where(s => s.UserId == id && s.Status == "active" && s.ExpiryDate > now)
                .OrderByDescending(s => s.ExpiryDate)
                .FirstOrDefaultAsync();

            AdminSubscriptionDto? subDto = null;
            if (sub != null)
            {
                subDto = new AdminSubscriptionDto
                {
                    PlanCode = sub.Plan.Code,
                    PlanName = sub.Plan.Name,
                    Status = sub.Status,
                    ExpiryDate = sub.ExpiryDate,
                };
            }
            else if (u.TrialStartDate.HasValue)
            {
                var trialEnd = u.TrialStartDate.Value.AddDays(7);
                var remaining = (int)Math.Ceiling((trialEnd - now).TotalDays);
                subDto = new AdminSubscriptionDto
                {
                    PlanCode = "TRIAL",
                    PlanName = "7天試用",
                    Status = remaining > 0 ? "trial" : "trial_expired",
                    ExpiryDate = trialEnd,
                    IsInTrial = remaining > 0,
                    TrialStartDate = u.TrialStartDate,
                    TrialDaysRemaining = remaining > 0 ? remaining : 0,
                };
            }

            return Ok(new AdminUserDto
            {
                Id = u.Id,
                Name = u.Name,
                Email = u.Email ?? string.Empty,
                Phone = u.Phone,
                PostalCode = u.PostalCode,
                Address = u.Address,
                TaxId = u.TaxId,
                Points = u.Points,
                BirthYear = u.BirthYear,
                BirthMonth = u.BirthMonth,
                BirthDay = u.BirthDay,
                BirthHour = u.BirthHour,
                BirthGender = u.BirthGender,
                Subscription = subDto,
            });
        }

        // PUT /api/Admin/users/{id}
        [HttpPut("users/{id}")]
        public async Task<IActionResult> UpdateUser(string id, [FromBody] UpdateUserRequest req)
        {
            if (!IsAdmin()) return Forbid();
            var user = await _userManager.FindByIdAsync(id);
            if (user == null) return NotFound();
            if (req.Name != null) user.Name = req.Name;
            if (req.Phone != null) user.Phone = req.Phone;
            if (req.PostalCode != null) user.PostalCode = req.PostalCode;
            if (req.Address != null) user.Address = req.Address;
            if (req.TaxId != null) user.TaxId = req.TaxId;
            var result = await _userManager.UpdateAsync(user);
            if (!result.Succeeded)
                return BadRequest(new { message = result.Errors.FirstOrDefault()?.Description });
            return Ok(new { message = "更新成功" });
        }

        // POST /api/Admin/users/{id}/change-password
        [HttpPost("users/{id}/change-password")]
        public async Task<IActionResult> ForceChangePassword(string id, [FromBody] ForceChangePasswordRequest req)
        {
            if (!IsAdmin()) return Forbid();
            if (string.IsNullOrEmpty(req.NewPassword) || req.NewPassword.Length < 6)
                return BadRequest(new { message = "密碼至少需要 6 位數" });
            var user = await _userManager.FindByIdAsync(id);
            if (user == null) return NotFound();
            var token = await _userManager.GeneratePasswordResetTokenAsync(user);
            var result = await _userManager.ResetPasswordAsync(user, token, req.NewPassword);
            if (!result.Succeeded)
                return BadRequest(new { message = result.Errors.FirstOrDefault()?.Description });
            return Ok(new { message = "密碼已強制變更" });
        }

        // POST /api/Admin/users/{id}/adjust-points
        [HttpPost("users/{id}/adjust-points")]
        public async Task<IActionResult> AdjustPoints(string id, [FromBody] AdjustPointsRequest req)
        {
            if (!IsAdmin()) return Forbid();
            var user = await _userManager.FindByIdAsync(id);
            if (user == null) return NotFound();
            if (user.Points + req.Amount < 0)
                return BadRequest(new { message = "點數不足，無法扣除" });
            user.Points += req.Amount;
            _context.PointRecords.Add(new PointRecord
            {
                UserId = user.Id,
                Amount = req.Amount,
                Description = $"管理員調整：{req.Reason}",
                CreatedAt = DateTime.UtcNow,
            });
            await _context.SaveChangesAsync();
            return Ok(new { message = "點數已調整", newPoints = user.Points });
        }

        // GET /api/Admin/atm-requests?status=pending
        [HttpGet("atm-requests")]
        public async Task<IActionResult> GetAtmRequests([FromQuery] string? status)
        {
            if (!IsAdmin()) return Forbid();
            var query = _context.AtmPaymentRequests.AsQueryable();
            if (!string.IsNullOrEmpty(status))
                query = query.Where(r => r.Status == status);
            var requests = await query
                .OrderByDescending(r => r.CreatedAt)
                .ToListAsync();
            var userIds = requests.Select(r => r.UserId).Distinct().ToList();
            var users = await _userManager.Users
                .Where(u => userIds.Contains(u.Id))
                .ToDictionaryAsync(u => u.Id, u => new { u.Email, u.Name });
            var result = requests.Select(r => new
            {
                r.Id,
                r.UserId,
                userEmail = users.ContainsKey(r.UserId) ? users[r.UserId].Email : "",
                userName  = users.ContainsKey(r.UserId) ? users[r.UserId].Name  : "",
                r.PackageId,
                r.Points,
                r.PriceTwd,
                r.TransferDate,
                r.AccountLast5,
                r.Status,
                r.AdminNote,
                r.CreatedAt,
                r.ProcessedAt,
            });
            return Ok(result);
        }

        // POST /api/Admin/atm-requests/{id}/approve
        [HttpPost("atm-requests/{id}/approve")]
        public async Task<IActionResult> ApproveAtmRequest(int id, [FromBody] AtmApproveRequest req)
        {
            if (!IsAdmin()) return Forbid();
            var request = await _context.AtmPaymentRequests.FindAsync(id);
            if (request == null) return NotFound();
            if (request.Status != "pending")
                return BadRequest(new { message = "此申請已處理" });
            var user = await _userManager.FindByIdAsync(request.UserId);
            if (user == null)
                return NotFound(new { message = "找不到用戶" });
            user.Points += request.Points;
            request.Status = "approved";
            request.AdminNote = req.Note;
            request.ProcessedAt = DateTime.UtcNow;
            _context.PointRecords.Add(new PointRecord
            {
                UserId = user.Id,
                Amount = request.Points,
                Description = $"ATM 轉帳儲值 {request.PackageId} 套餐",
                CreatedAt = DateTime.UtcNow,
            });
            await _context.SaveChangesAsync();
            return Ok(new { message = "已批准，點數入帳成功" });
        }

        // POST /api/Admin/atm-requests/{id}/reject
        [HttpPost("atm-requests/{id}/reject")]
        public async Task<IActionResult> RejectAtmRequest(int id, [FromBody] AtmRejectRequest req)
        {
            if (!IsAdmin()) return Forbid();
            var request = await _context.AtmPaymentRequests.FindAsync(id);
            if (request == null) return NotFound();
            if (request.Status != "pending")
                return BadRequest(new { message = "此申請已處理" });
            request.Status = "rejected";
            request.AdminNote = req.Note;
            request.ProcessedAt = DateTime.UtcNow;
            await _context.SaveChangesAsync();
            return Ok(new { message = "已拒絕" });
        }

        // ─── Product catalog management ───────────────────────────────────────

        // GET /api/Admin/products
        [HttpGet("products")]
        public async Task<IActionResult> GetProducts()
        {
            if (!IsAdmin()) return Forbid();
            var products = await _context.Products.OrderBy(p => p.SortOrder).ToListAsync();
            return Ok(products);
        }

        // PUT /api/Admin/products/{code}
        [HttpPut("products/{code}")]
        public async Task<IActionResult> UpdateProduct(string code, [FromBody] ProductUpdateRequest req)
        {
            if (!IsAdmin()) return Forbid();
            var product = await _context.Products.FirstOrDefaultAsync(p => p.Code == code);
            if (product == null) return NotFound();
            if (req.Name != null) product.Name = req.Name;
            if (req.PointCost.HasValue) product.PointCost = req.PointCost;
            if (req.PriceTwd.HasValue) product.PriceTwd = req.PriceTwd;
            if (req.IsActive.HasValue) product.IsActive = req.IsActive.Value;
            if (req.Description != null) product.Description = req.Description;
            if (req.SortOrder.HasValue) product.SortOrder = req.SortOrder.Value;
            product.UpdatedAt = DateTime.UtcNow;
            await _context.SaveChangesAsync();
            return Ok(new { message = "商品已更新", product });
        }

        // POST /api/Admin/products
        [HttpPost("products")]
        public async Task<IActionResult> CreateProduct([FromBody] Product product)
        {
            if (!IsAdmin()) return Forbid();
            if (await _context.Products.AnyAsync(p => p.Code == product.Code))
                return BadRequest(new { message = "商品代碼已存在" });
            product.UpdatedAt = DateTime.UtcNow;
            _context.Products.Add(product);
            await _context.SaveChangesAsync();
            return Ok(new { message = "商品已建立", product });
        }

        // ─── Membership plan management ───────────────────────────────────────

        // GET /api/Admin/plans
        [HttpGet("plans")]
        public async Task<IActionResult> GetPlans()
        {
            if (!IsAdmin()) return Forbid();
            var plans = await _context.MembershipPlans
                .OrderBy(p => p.SortOrder)
                .Include(p => p.Benefits)
                .ToListAsync();
            return Ok(plans);
        }

        // PUT /api/Admin/plans/{code}
        [HttpPut("plans/{code}")]
        public async Task<IActionResult> UpdatePlan(string code, [FromBody] PlanUpdateRequest req)
        {
            if (!IsAdmin()) return Forbid();
            var plan = await _context.MembershipPlans.FirstOrDefaultAsync(p => p.Code == code);
            if (plan == null) return NotFound();
            if (req.Name != null) plan.Name = req.Name;
            if (req.PriceTwd.HasValue) plan.PriceTwd = req.PriceTwd.Value;
            if (req.IsActive.HasValue) plan.IsActive = req.IsActive.Value;
            if (req.Description != null) plan.Description = req.Description;
            await _context.SaveChangesAsync();
            return Ok(new { message = "套餐已更新", plan });
        }

        // POST /api/Admin/plans/{code}/benefits
        [HttpPost("plans/{code}/benefits")]
        public async Task<IActionResult> AddPlanBenefit(string code, [FromBody] MembershipPlanBenefit benefit)
        {
            if (!IsAdmin()) return Forbid();
            var plan = await _context.MembershipPlans.FirstOrDefaultAsync(p => p.Code == code);
            if (plan == null) return NotFound();
            benefit.PlanId = plan.Id;
            _context.MembershipPlanBenefits.Add(benefit);
            await _context.SaveChangesAsync();
            return Ok(new { message = "福利已新增", benefit });
        }

        // DELETE /api/Admin/plans/benefits/{id}
        [HttpDelete("plans/benefits/{id}")]
        public async Task<IActionResult> DeletePlanBenefit(int id)
        {
            if (!IsAdmin()) return Forbid();
            var benefit = await _context.MembershipPlanBenefits.FindAsync(id);
            if (benefit == null) return NotFound();
            _context.MembershipPlanBenefits.Remove(benefit);
            await _context.SaveChangesAsync();
            return Ok(new { message = "福利已刪除" });
        }

        // ─── Subscription management ──────────────────────────────────────────

        // POST /api/Admin/users/{id}/grant-subscription
        [HttpPost("users/{id}/grant-subscription")]
        public async Task<IActionResult> GrantSubscription(string id, [FromBody] GrantSubscriptionRequest req)
        {
            if (!IsAdmin()) return Forbid();
            var user = await _userManager.FindByIdAsync(id);
            if (user == null) return NotFound();
            var plan = await _context.MembershipPlans.FirstOrDefaultAsync(p => p.Code == req.PlanCode && p.IsActive);
            if (plan == null) return BadRequest(new { message = "套餐不存在" });
            var now = DateTime.UtcNow;
            _context.UserSubscriptions.Add(new UserSubscription
            {
                UserId     = id,
                PlanId     = plan.Id,
                StartDate  = now,
                ExpiryDate = now.AddDays(plan.DurationDays),
                Status     = "active",
                PaymentRef = req.Note ?? "admin-grant",
                CreatedAt  = now
            });
            await _context.SaveChangesAsync();
            return Ok(new { message = $"已為用戶開通 {plan.Name}" });
        }

        // ─── Booking request management ───────────────────────────────────────

        // GET /api/Admin/bookings?type=&status=&page=1
        [HttpGet("bookings")]
        public async Task<IActionResult> GetBookings([FromQuery] string? type, [FromQuery] string? status, [FromQuery] int page = 1)
        {
            if (!IsAdmin()) return Forbid();
            var query = _context.BookingRequests.AsQueryable();
            if (!string.IsNullOrWhiteSpace(type)) query = query.Where(b => b.ServiceType == type);
            if (!string.IsNullOrWhiteSpace(status)) query = query.Where(b => b.Status == status);
            var total = await query.CountAsync();
            var items = await query.OrderByDescending(b => b.CreatedAt).Skip((page - 1) * 20).Take(20).ToListAsync();
            return Ok(new { total, page, items });
        }

        // PUT /api/Admin/bookings/{id}/status
        [HttpPut("bookings/{id}/status")]
        public async Task<IActionResult> UpdateBookingStatus(int id, [FromBody] BookingStatusUpdateRequest req)
        {
            if (!IsAdmin()) return Forbid();
            var booking = await _context.BookingRequests.FindAsync(id);
            if (booking == null) return NotFound();
            if (req.Status != null) booking.Status = req.Status;
            if (req.AdminNote != null) booking.AdminNote = req.AdminNote;
            await _context.SaveChangesAsync();
            return Ok(new { message = "預約狀態已更新", booking });
        }
    }

    public class ProductUpdateRequest
    {
        public string? Name { get; set; }
        public int? PointCost { get; set; }
        public int? PriceTwd { get; set; }
        public bool? IsActive { get; set; }
        public string? Description { get; set; }
        public int? SortOrder { get; set; }
    }

    public class PlanUpdateRequest
    {
        public string? Name { get; set; }
        public int? PriceTwd { get; set; }
        public bool? IsActive { get; set; }
        public string? Description { get; set; }
    }

    public class GrantSubscriptionRequest
    {
        public string PlanCode { get; set; } = "";
        public string? Note { get; set; }
    }

    public class BookingStatusUpdateRequest
    {
        public string? Status { get; set; }   // pending / confirmed / completed / cancelled
        public string? AdminNote { get; set; }
    }
}
