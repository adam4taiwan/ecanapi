using Ecanapi.Data;
using Ecanapi.Models;
using Ecanapi.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Security.Claims;

namespace Ecanapi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ReportsController : ControllerBase
    {
        private readonly ApplicationDbContext _context;
        private readonly IConfiguration _config;
        private readonly IEmailService _email;
        private readonly UserManager<ApplicationUser> _userManager;

        public ReportsController(ApplicationDbContext context, IConfiguration config, IEmailService email, UserManager<ApplicationUser> userManager)
        {
            _context = context;
            _config = config;
            _email = email;
            _userManager = userManager;
        }

        private string? GetCurrentUserIdentity()
        {
            return User.FindFirstValue(ClaimTypes.Email)
                ?? User.FindFirstValue(ClaimTypes.Name)
                ?? User.FindFirst("unique_name")?.Value;
        }

        private bool IsAdmin()
        {
            var identity = GetCurrentUserIdentity();
            if (string.IsNullOrEmpty(identity)) return false;
            var adminEmail = _config["Admin:Email"];
            return !string.IsNullOrEmpty(adminEmail) &&
                   string.Equals(identity, adminEmail, StringComparison.OrdinalIgnoreCase);
        }

        // ─── User endpoints ───────────────────────────────────────────────

        // GET /api/Reports/my
        [HttpGet("my")]
        [Authorize]
        public async Task<IActionResult> GetMyReports()
        {
            var identity = GetCurrentUserIdentity();
            if (string.IsNullOrEmpty(identity)) return Unauthorized();

            var user = await _userManager.FindByEmailAsync(identity)
                    ?? await _userManager.FindByNameAsync(identity);
            if (user == null) return Unauthorized();

            var reports = await _context.UserReports
                .Where(r => r.UserId == user.Id)
                .OrderByDescending(r => r.CreatedAt)
                .Select(r => new
                {
                    r.Id,
                    r.ReportType,
                    r.Title,
                    r.Status,
                    r.CreatedAt,
                    r.ApprovedAt,
                    r.AdminNote,
                    HasDownloadToken = r.DownloadToken != null && r.DownloadTokenExpiry > DateTime.UtcNow
                })
                .ToListAsync();

            return Ok(reports);
        }

        // GET /api/Reports/by-token?token={token}
        // Public: return report content for DOCX generation on frontend
        [HttpGet("by-token")]
        public async Task<IActionResult> GetByToken([FromQuery] string token)
        {
            if (string.IsNullOrEmpty(token)) return BadRequest(new { error = "無效的下載連結" });

            var report = await _context.UserReports
                .FirstOrDefaultAsync(r => r.DownloadToken == token);

            if (report == null) return NotFound(new { error = "找不到命書，請確認連結正確" });
            if (report.Status != "approved") return BadRequest(new { error = "命書尚未審核完成" });
            if (report.DownloadTokenExpiry == null || report.DownloadTokenExpiry < DateTime.UtcNow)
                return BadRequest(new { error = "下載連結已過期，請聯繫玉洞子重新發送" });

            var user = await _userManager.FindByIdAsync(report.UserId);
            var personName = user?.UserName ?? "命主";

            var (bookTitle, skipTitle) = report.ReportType switch
            {
                "bazi" or "bazi-ziwei" => ("八 字 命 書", "八 字 命 書"),
                "daiyun" => ("大 運 命 書", "大 運 命 書"),
                "liunian" => ("流 年 命 書", "流 年 命 書"),
                "lifelong" => ("終 身 命 書", "終 身 命 書"),
                _ => ("命書", "命書")
            };

            return Ok(new
            {
                reportId = report.Id,
                title = report.Title,
                reportType = report.ReportType,
                content = report.Content,
                personName,
                bookTitle,
                skipTitle
            });
        }

        // ─── Admin endpoints ──────────────────────────────────────────────

        // GET /api/Reports/admin/list?status=pending_review|approved|rejected|all
        [HttpGet("admin/list")]
        [Authorize]
        public async Task<IActionResult> GetReportsAdmin([FromQuery] string? status = "pending_review")
        {
            if (!IsAdmin()) return Forbid();

            var query = _context.UserReports.AsQueryable();
            if (status != "all")
                query = query.Where(r => r.Status == (status ?? "pending_review"));
            query = query.OrderByDescending(r => r.CreatedAt);

            var reports = await query
                .Select(r => new
                {
                    r.Id,
                    r.UserId,
                    r.ReportType,
                    r.Title,
                    r.Status,
                    r.CreatedAt,
                    r.ApprovedAt,
                    r.AdminNote,
                    ContentPreview = r.Content.Length > 300 ? r.Content.Substring(0, 300) + "..." : r.Content
                })
                .ToListAsync();

            var userIds = reports.Select(r => r.UserId).Distinct().ToList();
            var users = await _context.Users
                .Where(u => userIds.Contains(u.Id))
                .Select(u => new { u.Id, u.Email, u.UserName })
                .ToListAsync();
            var userMap = users.ToDictionary(u => u.Id);

            var result = reports.Select(r =>
            {
                userMap.TryGetValue(r.UserId, out var u);
                return new
                {
                    r.Id,
                    r.UserId,
                    r.ReportType,
                    r.Title,
                    r.Status,
                    r.CreatedAt,
                    r.ApprovedAt,
                    r.AdminNote,
                    r.ContentPreview,
                    UserEmail = u?.Email,
                    UserName = u?.UserName
                };
            });

            return Ok(result);
        }

        // GET /api/Reports/admin/{id}
        [HttpGet("admin/{id:guid}")]
        [Authorize]
        public async Task<IActionResult> GetReportAdmin(Guid id)
        {
            if (!IsAdmin()) return Forbid();

            var report = await _context.UserReports.FirstOrDefaultAsync(r => r.Id == id);
            if (report == null) return NotFound();

            var user = await _userManager.FindByIdAsync(report.UserId);

            return Ok(new
            {
                report.Id,
                report.UserId,
                report.ReportType,
                report.Title,
                report.Status,
                report.Content,
                report.Parameters,
                report.CreatedAt,
                report.ApprovedAt,
                report.AdminNote,
                UserEmail = user?.Email,
                UserName = user?.UserName
            });
        }

        // POST /api/Reports/admin/{id}/approve
        [HttpPost("admin/{id:guid}/approve")]
        [Authorize]
        public async Task<IActionResult> ApproveReport(Guid id)
        {
            if (!IsAdmin()) return Forbid();

            var report = await _context.UserReports.FirstOrDefaultAsync(r => r.Id == id);
            if (report == null) return NotFound(new { error = "找不到命書" });
            if (report.Status == "approved") return BadRequest(new { error = "命書已核准" });

            var user = await _userManager.FindByIdAsync(report.UserId);
            if (user == null) return BadRequest(new { error = "找不到用戶" });

            var token = Guid.NewGuid().ToString("N");
            var expiry = DateTime.UtcNow.AddHours(72);

            report.Status = "approved";
            report.ApprovedAt = DateTime.UtcNow;
            report.DownloadToken = token;
            report.DownloadTokenExpiry = expiry;
            await _context.SaveChangesAsync();

            var frontendBase = _config["App:FrontendUrl"] ?? "https://yudongzi.tw";
            var downloadUrl = $"{frontendBase}/disk?downloadToken={token}";
            var expiryDesc = TimeZoneInfo.ConvertTimeFromUtc(expiry,
                TimeZoneInfo.FindSystemTimeZoneById("Asia/Taipei")).ToString("yyyy/MM/dd HH:mm");

            var toEmail = user.Email ?? string.Empty;
            var toName = user.UserName ?? "命主";
            await _email.SendReportReadyAsync(toEmail, toName, report.Title, downloadUrl, expiryDesc);

            return Ok(new { message = "命書已核准，Email 已發送", downloadUrl, expiryAt = expiry });
        }

        // POST /api/Reports/admin/{id}/reject
        [HttpPost("admin/{id:guid}/reject")]
        [Authorize]
        public async Task<IActionResult> RejectReport(Guid id, [FromBody] RejectReportRequest req)
        {
            if (!IsAdmin()) return Forbid();

            var report = await _context.UserReports.FirstOrDefaultAsync(r => r.Id == id);
            if (report == null) return NotFound(new { error = "找不到命書" });

            report.Status = "rejected";
            report.AdminNote = req.Note;
            await _context.SaveChangesAsync();

            return Ok(new { message = "已退回" });
        }

        // POST /api/Reports/admin/{id}/resend
        [HttpPost("admin/{id:guid}/resend")]
        [Authorize]
        public async Task<IActionResult> ResendReport(Guid id)
        {
            if (!IsAdmin()) return Forbid();

            var report = await _context.UserReports.FirstOrDefaultAsync(r => r.Id == id);
            if (report == null) return NotFound(new { error = "找不到命書" });
            if (report.Status != "approved") return BadRequest(new { error = "只有已核准的命書才能重發" });

            var user = await _userManager.FindByIdAsync(report.UserId);
            if (user == null) return BadRequest(new { error = "找不到用戶" });

            var token = Guid.NewGuid().ToString("N");
            var expiry = DateTime.UtcNow.AddHours(72);
            report.DownloadToken = token;
            report.DownloadTokenExpiry = expiry;
            await _context.SaveChangesAsync();

            var frontendBase = _config["App:FrontendUrl"] ?? "https://yudongzi.tw";
            var downloadUrl = $"{frontendBase}/disk?downloadToken={token}";
            var expiryDesc = TimeZoneInfo.ConvertTimeFromUtc(expiry,
                TimeZoneInfo.FindSystemTimeZoneById("Asia/Taipei")).ToString("yyyy/MM/dd HH:mm");

            var toEmail = user.Email ?? string.Empty;
            var toName = user.UserName ?? "命主";
            await _email.SendReportReadyAsync(toEmail, toName, report.Title, downloadUrl, expiryDesc);

            return Ok(new { message = "Email 已重新發送", downloadUrl, expiryAt = expiry });
        }
    }

    public class RejectReportRequest
    {
        public string Note { get; set; } = string.Empty;
    }
}
