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
    }
}
