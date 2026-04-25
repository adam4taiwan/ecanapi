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
            var user = await _userManager.FindByEmailAsync(identity) ?? await _userManager.FindByNameAsync(identity);
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
                    HasDownloadToken = r.DownloadToken != null && r.DownloadTokenExpiry > DateTime.UtcNow,
                    HasApprovedDocx = r.ApprovedDocxBytes != null
                })
                .ToListAsync();

            return Ok(reports);
        }

        // GET /api/Reports/download-approved?token={token}
        // Public direct download of the admin-approved DOCX file
        [HttpGet("download-approved")]
        public async Task<IActionResult> DownloadApproved([FromQuery] string token)
        {
            if (string.IsNullOrEmpty(token))
                return BadRequest("無效的下載連結");

            var report = await _context.UserReports.FirstOrDefaultAsync(r => r.DownloadToken == token);
            if (report == null) return NotFound("找不到命書，請確認連結正確");
            if (report.Status != "approved") return BadRequest("命書尚未審核完成");
            if (report.DownloadTokenExpiry == null || report.DownloadTokenExpiry < DateTime.UtcNow)
                return BadRequest("下載連結已過期，請聯繫玉洞子重新發送");
            if (report.ApprovedDocxBytes == null || report.ApprovedDocxBytes.Length == 0)
                return StatusCode(500, "命書檔案尚未上傳，請聯繫玉洞子");

            var user = await _userManager.FindByIdAsync(report.UserId);
            var personName = !string.IsNullOrEmpty(user?.ChartName) ? user.ChartName : (user?.UserName ?? "命主");
            var fileName = report.ApprovedDocxFileName
                ?? $"{personName}_{report.Title}.docx";

            return File(report.ApprovedDocxBytes,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                fileName);
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
                    HasApprovedDocx = r.ApprovedDocxBytes != null,
                    ApprovedDocxFileName = r.ApprovedDocxFileName,
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
                    r.HasApprovedDocx,
                    r.ApprovedDocxFileName,
                    r.ContentPreview,
                    UserEmail = u?.Email,
                    UserName = u?.UserName
                };
            });

            return Ok(result);
        }

        // GET /api/Reports/admin/{id}
        // Full report content for admin review
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
                HasApprovedDocx = report.ApprovedDocxBytes != null,
                report.ApprovedDocxFileName,
                UserEmail = user?.Email,
                UserName = user?.UserName
            });
        }

        // GET /api/Reports/admin/{id}/download-draft-docx
        // Admin downloads AI-generated DOCX draft (for local editing)
        [HttpGet("admin/{id:guid}/download-draft-docx")]
        [Authorize]
        public async Task<IActionResult> DownloadDraftDocx(Guid id)
        {
            if (!IsAdmin()) return Forbid();

            var report = await _context.UserReports.FirstOrDefaultAsync(r => r.Id == id);
            if (report == null) return NotFound(new { error = "找不到命書" });

            // If admin-approved DOCX already exists, serve that instead
            if (report.ApprovedDocxBytes != null && report.ApprovedDocxBytes.Length > 0)
            {
                var approvedName = report.ApprovedDocxFileName ?? $"修正版_{report.Title}.docx";
                return File(report.ApprovedDocxBytes,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    approvedName);
            }

            // Build DOCX from AI-generated content using the export endpoint
            // Return the report content as JSON for the admin panel to call export-generic-docx
            var user = await _userManager.FindByIdAsync(report.UserId);
            var personName = !string.IsNullOrEmpty(user?.ChartName) ? user.ChartName : (user?.UserName ?? "命主");

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

        // POST /api/Reports/admin/{id}/upload-docx
        // Admin uploads the corrected DOCX file
        [HttpPost("admin/{id:guid}/upload-docx")]
        [Authorize]
        [RequestSizeLimit(20 * 1024 * 1024)] // 20 MB
        public async Task<IActionResult> UploadDocx(Guid id, IFormFile file)
        {
            if (!IsAdmin()) return Forbid();
            if (file == null || file.Length == 0) return BadRequest(new { error = "請選擇檔案" });
            if (!file.FileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                return BadRequest(new { error = "只接受 .docx 格式" });

            var report = await _context.UserReports.FirstOrDefaultAsync(r => r.Id == id);
            if (report == null) return NotFound(new { error = "找不到命書" });

            using var ms = new MemoryStream();
            await file.CopyToAsync(ms);
            report.ApprovedDocxBytes = ms.ToArray();
            report.ApprovedDocxFileName = file.FileName;
            await _context.SaveChangesAsync();

            return Ok(new { message = $"已上傳：{file.FileName}（{ms.Length / 1024} KB）" });
        }

        // POST /api/Reports/admin/{id}/approve
        // Approve + generate download token + send email to user
        [HttpPost("admin/{id:guid}/approve")]
        [Authorize]
        public async Task<IActionResult> ApproveReport(Guid id)
        {
            if (!IsAdmin()) return Forbid();

            var report = await _context.UserReports.FirstOrDefaultAsync(r => r.Id == id);
            if (report == null) return NotFound(new { error = "找不到命書" });
            if (report.Status == "approved") return BadRequest(new { error = "命書已核准" });
            if (report.ApprovedDocxBytes == null || report.ApprovedDocxBytes.Length == 0)
                return BadRequest(new { error = "請先上傳修正後的命書 DOCX，再執行核准" });

            var user = await _userManager.FindByIdAsync(report.UserId);
            if (user == null) return BadRequest(new { error = "找不到用戶" });

            var token = Guid.NewGuid().ToString("N");
            var expiry = DateTime.UtcNow.AddHours(72);

            report.Status = "approved";
            report.ApprovedAt = DateTime.UtcNow;
            report.DownloadToken = token;
            report.DownloadTokenExpiry = expiry;
            await _context.SaveChangesAsync();

            // Email link points directly to Ecanapi download endpoint
            var apiBase = _config["ECPay:ApiBase"] ?? "https://ecanapi.fly.dev";
            var downloadUrl = $"{apiBase}/api/Reports/download-approved?token={token}";

            string expiryDesc;
            try
            {
                expiryDesc = TimeZoneInfo.ConvertTimeFromUtc(expiry,
                    TimeZoneInfo.FindSystemTimeZoneById("Asia/Taipei")).ToString("yyyy/MM/dd HH:mm");
            }
            catch { expiryDesc = expiry.AddHours(8).ToString("yyyy/MM/dd HH:mm"); }

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
        // Resend email with new token (if 72hr expired)
        [HttpPost("admin/{id:guid}/resend")]
        [Authorize]
        public async Task<IActionResult> ResendReport(Guid id)
        {
            if (!IsAdmin()) return Forbid();
            var report = await _context.UserReports.FirstOrDefaultAsync(r => r.Id == id);
            if (report == null) return NotFound(new { error = "找不到命書" });
            if (report.Status != "approved") return BadRequest(new { error = "只有已核准的命書才能重發" });
            if (report.ApprovedDocxBytes == null) return BadRequest(new { error = "尚無上傳的命書檔案" });

            var user = await _userManager.FindByIdAsync(report.UserId);
            if (user == null) return BadRequest(new { error = "找不到用戶" });

            var token = Guid.NewGuid().ToString("N");
            var expiry = DateTime.UtcNow.AddHours(72);
            report.DownloadToken = token;
            report.DownloadTokenExpiry = expiry;
            await _context.SaveChangesAsync();

            var apiBase = _config["ECPay:ApiBase"] ?? "https://ecanapi.fly.dev";
            var downloadUrl = $"{apiBase}/api/Reports/download-approved?token={token}";

            string expiryDesc;
            try
            {
                expiryDesc = TimeZoneInfo.ConvertTimeFromUtc(expiry,
                    TimeZoneInfo.FindSystemTimeZoneById("Asia/Taipei")).ToString("yyyy/MM/dd HH:mm");
            }
            catch { expiryDesc = expiry.AddHours(8).ToString("yyyy/MM/dd HH:mm"); }

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
