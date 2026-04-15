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
    public class BlessingBookingRequest
    {
        public string ServiceCode { get; set; } = "";  // BLESSING_ANTAISUI / BLESSING_LIGHT / BLESSING_WEALTH / BLESSING_PRAYER
        public string Name { get; set; } = "";
        public string BirthDate { get; set; } = "";
        public bool IsLunar { get; set; } = false;
        public string ContactType { get; set; } = "";  // line / wechat / phone
        public string ContactInfo { get; set; } = "";
        public string? Notes { get; set; }
    }

    public class ConsultationBookingRequest
    {
        public string Name { get; set; } = "";
        public string BirthDate { get; set; } = "";
        public bool IsLunar { get; set; } = false;
        public string Topic { get; set; } = "";        // 婚姻 / 事業 / 財運 / 健康 / 其他
        public string ContactType { get; set; } = "";  // line / wechat / phone
        public string ContactInfo { get; set; } = "";
        public string? PreferredDate { get; set; }
        public string? Notes { get; set; }
    }

    public class UpdateBookingStatusRequest
    {
        public string Status { get; set; } = "";       // pending / confirmed / completed / cancelled
        public string? AdminNote { get; set; }
    }

    [ApiController]
    [Route("api/[controller]")]
    public class BookingController : ControllerBase
    {
        private readonly ApplicationDbContext _db;
        private readonly IEmailService _email;
        private readonly UserManager<ApplicationUser> _userManager;
        private readonly IConfiguration _config;

        private static readonly HashSet<string> ValidBlessingCodes = new()
        {
            "BLESSING_ANTAISUI", "BLESSING_LIGHT", "BLESSING_WEALTH", "BLESSING_PRAYER"
        };

        public BookingController(ApplicationDbContext db, IEmailService email, UserManager<ApplicationUser> userManager, IConfiguration config)
        {
            _db = db;
            _email = email;
            _userManager = userManager;
            _config = config;
        }

        private bool IsAdmin()
        {
            var email = User.FindFirstValue(ClaimTypes.Email);
            var adminEmail = _config["Admin:Email"];
            return !string.IsNullOrEmpty(email) && email == adminEmail;
        }

        // POST /api/Booking/blessing
        [HttpPost("blessing")]
        [Authorize]
        public async Task<IActionResult> SubmitBlessing([FromBody] BlessingBookingRequest req)
        {
            if (string.IsNullOrWhiteSpace(req.ServiceCode) || !ValidBlessingCodes.Contains(req.ServiceCode))
                return BadRequest(new { message = "無效的服務代碼" });
            if (string.IsNullOrWhiteSpace(req.Name))
                return BadRequest(new { message = "請填寫姓名" });
            if (string.IsNullOrWhiteSpace(req.ContactInfo))
                return BadRequest(new { message = "請填寫聯絡方式" });

            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            var user = userId != null ? await _userManager.FindByIdAsync(userId) : null;

            var booking = new BookingRequest
            {
                UserId = userId,
                ServiceType = "blessing",
                ServiceCode = req.ServiceCode,
                Name = req.Name,
                BirthDate = req.BirthDate,
                IsLunar = req.IsLunar,
                ContactType = req.ContactType,
                ContactInfo = req.ContactInfo,
                Notes = req.Notes,
                Status = "pending",
                CreatedAt = DateTime.UtcNow,
            };

            _db.BookingRequests.Add(booking);
            await _db.SaveChangesAsync();

            // Send emails (fire & forget, don't block response)
            _ = _email.SendAdminBookingNotifyAsync("blessing", req.Name, req.ContactType, req.ContactInfo, req.Notes, null, user?.Email, booking.Id);
            if (!string.IsNullOrEmpty(user?.Email))
                _ = _email.SendClientConfirmationAsync(user.Email, req.Name, "blessing");

            return Ok(new { message = "登記成功，我們將於 1-2 個工作日內與您聯繫確認。", bookingId = booking.Id });
        }

        // POST /api/Booking/consultation
        [HttpPost("consultation")]
        [Authorize]
        public async Task<IActionResult> SubmitConsultation([FromBody] ConsultationBookingRequest req)
        {
            if (string.IsNullOrWhiteSpace(req.Name))
                return BadRequest(new { message = "請填寫姓名" });
            if (string.IsNullOrWhiteSpace(req.ContactInfo))
                return BadRequest(new { message = "請填寫聯絡方式" });
            if (string.IsNullOrWhiteSpace(req.Topic))
                return BadRequest(new { message = "請選擇諮詢主題" });

            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            var user = userId != null ? await _userManager.FindByIdAsync(userId) : null;

            var notes = string.IsNullOrWhiteSpace(req.Notes) ? req.Topic : $"[{req.Topic}] {req.Notes}";

            var booking = new BookingRequest
            {
                UserId = userId,
                ServiceType = "consultation",
                ServiceCode = "CONSULT_VIDEO",
                Name = req.Name,
                BirthDate = req.BirthDate,
                IsLunar = req.IsLunar,
                ContactType = req.ContactType,
                ContactInfo = req.ContactInfo,
                Notes = notes,
                PreferredDate = req.PreferredDate,
                Status = "pending",
                CreatedAt = DateTime.UtcNow,
            };

            _db.BookingRequests.Add(booking);
            await _db.SaveChangesAsync();

            // Send emails
            _ = _email.SendAdminBookingNotifyAsync("consultation", req.Name, req.ContactType, req.ContactInfo, notes, req.PreferredDate, user?.Email, booking.Id);
            if (!string.IsNullOrEmpty(user?.Email))
                _ = _email.SendClientConfirmationAsync(user.Email, req.Name, "consultation");

            return Ok(new { message = "預約申請已送出，玉洞子將於 1-2 個工作日內透過您指定的方式聯繫。", bookingId = booking.Id });
        }

        // GET /api/Booking/list?status=pending&type=&page=1
        [HttpGet("list")]
        [Authorize]
        public async Task<IActionResult> ListBookings([FromQuery] string? status, [FromQuery] string? type, [FromQuery] int page = 1)
        {
            if (!IsAdmin()) return Forbid();

            var query = _db.BookingRequests.AsQueryable();

            if (!string.IsNullOrEmpty(status))
                query = query.Where(b => b.Status == status);
            if (!string.IsNullOrEmpty(type))
                query = query.Where(b => b.ServiceType == type);

            var total = await query.CountAsync();
            var pageSize = 20;
            var items = await query
                .OrderByDescending(b => b.CreatedAt)
                .Skip((page - 1) * pageSize)
                .Take(pageSize)
                .ToListAsync();

            // Get user emails for display
            var userIds = items.Where(b => b.UserId != null).Select(b => b.UserId!).Distinct().ToList();
            var userEmails = await _db.Users
                .Where(u => userIds.Contains(u.Id))
                .ToDictionaryAsync(u => u.Id, u => u.Email ?? "");

            var result = items.Select(b => new
            {
                b.Id,
                b.ServiceType,
                b.ServiceCode,
                b.Name,
                b.ContactType,
                b.ContactInfo,
                b.BirthDate,
                b.IsLunar,
                b.Notes,
                b.PreferredDate,
                b.Status,
                b.AdminNote,
                b.CreatedAt,
                UserEmail = b.UserId != null && userEmails.ContainsKey(b.UserId) ? userEmails[b.UserId] : null,
            });

            return Ok(new { total, page, pageSize, items = result });
        }

        // PUT /api/Booking/{id}/status
        [HttpPut("{id}/status")]
        [Authorize]
        public async Task<IActionResult> UpdateStatus(int id, [FromBody] UpdateBookingStatusRequest req)
        {
            if (!IsAdmin()) return Forbid();

            var validStatuses = new[] { "pending", "confirmed", "completed", "cancelled" };
            if (!validStatuses.Contains(req.Status))
                return BadRequest(new { message = "無效的狀態" });

            var booking = await _db.BookingRequests.FindAsync(id);
            if (booking == null) return NotFound();

            booking.Status = req.Status;
            if (req.AdminNote != null)
                booking.AdminNote = req.AdminNote;

            await _db.SaveChangesAsync();
            return Ok(new { message = "更新成功" });
        }
    }
}
