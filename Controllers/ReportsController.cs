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
    public class ReportsController : ControllerBase
    {
        private readonly ApplicationDbContext _context;
        private readonly IConfiguration _config;

        public ReportsController(ApplicationDbContext context, IConfiguration config)
        {
            _context = context;
            _config = config;
        }

        // GET /api/Reports/my — list current user's reports (no content, just metadata)
        [Authorize]
        [HttpGet("my")]
        public async Task<IActionResult> GetMyReports()
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (string.IsNullOrEmpty(userId)) return Unauthorized();

            var reports = await _context.UserReports
                .Where(r => r.UserId == userId)
                .OrderByDescending(r => r.CreatedAt)
                .Select(r => new
                {
                    r.Id,
                    r.ReportType,
                    r.Title,
                    r.CreatedAt
                })
                .ToListAsync();

            return Ok(reports);
        }

        // GET /api/Reports/my/{id} — get full report content for current user
        [Authorize]
        [HttpGet("my/{id:guid}")]
        public async Task<IActionResult> GetMyReport(Guid id)
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (string.IsNullOrEmpty(userId)) return Unauthorized();

            var report = await _context.UserReports.FirstOrDefaultAsync(r => r.Id == id && r.UserId == userId);
            if (report == null) return NotFound(new { error = "命書記錄不存在" });

            return Ok(new
            {
                report.Id,
                report.ReportType,
                report.Title,
                report.Content,
                report.Parameters,
                report.CreatedAt
            });
        }

        // GET /api/Reports/admin — admin: list all reports
        [Authorize]
        [HttpGet("admin")]
        public async Task<IActionResult> AdminListReports([FromQuery] string? userId = null, [FromQuery] string? reportType = null, [FromQuery] int page = 1, [FromQuery] int pageSize = 50)
        {
            var requesterId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            var requester = await _context.Users.FindAsync(requesterId);
            var adminEmail = _config["Admin:Email"];
            if (requester == null || !string.Equals(requester.Email, adminEmail, StringComparison.OrdinalIgnoreCase))
                return StatusCode(403, new { error = "僅限管理員" });

            var query = _context.UserReports.AsQueryable();
            if (!string.IsNullOrEmpty(userId)) query = query.Where(r => r.UserId == userId);
            if (!string.IsNullOrEmpty(reportType)) query = query.Where(r => r.ReportType == reportType);

            int total = await query.CountAsync();
            var reports = await query
                .OrderByDescending(r => r.CreatedAt)
                .Skip((page - 1) * pageSize)
                .Take(pageSize)
                .Select(r => new { r.Id, r.UserId, r.ReportType, r.Title, r.CreatedAt })
                .ToListAsync();

            return Ok(new { total, page, pageSize, reports });
        }

        // DELETE /api/Reports/admin/{id} — admin: delete a report
        [Authorize]
        [HttpDelete("admin/{id:guid}")]
        public async Task<IActionResult> AdminDeleteReport(Guid id)
        {
            var requesterId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            var requester = await _context.Users.FindAsync(requesterId);
            var adminEmail = _config["Admin:Email"];
            if (requester == null || !string.Equals(requester.Email, adminEmail, StringComparison.OrdinalIgnoreCase))
                return StatusCode(403, new { error = "僅限管理員" });

            var report = await _context.UserReports.FindAsync(id);
            if (report == null) return NotFound(new { error = "記錄不存在" });

            _context.UserReports.Remove(report);
            await _context.SaveChangesAsync();
            return Ok(new { message = "已刪除" });
        }
    }
}
