using Ecanapi.Data;
using Ecanapi.Models;
using Microsoft.AspNetCore.Authorization;
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

    [ApiController]
    [Route("api/[controller]")]
    public class BookingController : ControllerBase
    {
        private readonly ApplicationDbContext _db;

        private static readonly HashSet<string> ValidBlessingCodes = new()
        {
            "BLESSING_ANTAISUI", "BLESSING_LIGHT", "BLESSING_WEALTH", "BLESSING_PRAYER"
        };

        // topics are free-form labels from frontend, no strict whitelist needed

        public BookingController(ApplicationDbContext db)
        {
            _db = db;
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
                Notes = string.IsNullOrWhiteSpace(req.Notes) ? req.Topic : $"[{req.Topic}] {req.Notes}",
                PreferredDate = req.PreferredDate,
                Status = "pending",
                CreatedAt = DateTime.UtcNow,
            };

            _db.BookingRequests.Add(booking);
            await _db.SaveChangesAsync();

            return Ok(new { message = "預約申請已送出，玉洞子將於 1-2 個工作日內透過您指定的方式聯繫。", bookingId = booking.Id });
        }
    }
}
