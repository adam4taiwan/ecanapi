using Ecanapi.Data;
using Ecanapi.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Security.Claims;

namespace Ecanapi.Controllers
{
    public class AddStudentRequest
    {
        public string UserId { get; set; } = string.Empty;
        public string? Note { get; set; }
    }

    [ApiController]
    [Route("api/[controller]")]
    [Authorize]
    public class StudentController : ControllerBase
    {
        private readonly ApplicationDbContext _context;
        private readonly UserManager<ApplicationUser> _userManager;
        private readonly IConfiguration _config;

        public StudentController(
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

        // GET /api/Student/check-access - check if logged-in user is in student whitelist
        [HttpGet("check-access")]
        public async Task<IActionResult> CheckAccess()
        {
            var email = User.FindFirstValue(ClaimTypes.Email);
            if (string.IsNullOrEmpty(email))
                return Unauthorized(new { hasAccess = false });

            var exists = await _context.StudentWhiteLists
                .AnyAsync(s => s.Email.ToLower() == email.ToLower());

            return Ok(new { hasAccess = exists, email });
        }

        // GET /api/Student/whitelist - list all whitelist entries (admin only)
        [HttpGet("whitelist")]
        public async Task<IActionResult> GetWhiteList()
        {
            if (!IsAdmin()) return Forbid();

            var list = await _context.StudentWhiteLists
                .OrderByDescending(s => s.AddedAt)
                .Select(s => new
                {
                    s.Id,
                    s.Email,
                    s.Note,
                    s.AddedByEmail,
                    addedAt = s.AddedAt.ToString("yyyy-MM-dd HH:mm")
                })
                .ToListAsync();

            return Ok(list);
        }

        // GET /api/Student/members - list all registered users for admin to select (admin only)
        [HttpGet("members")]
        public async Task<IActionResult> GetMembers([FromQuery] string? search)
        {
            if (!IsAdmin()) return Forbid();

            var query = _userManager.Users.AsQueryable();
            if (!string.IsNullOrEmpty(search))
                query = query.Where(u =>
                    (u.Email != null && u.Email.Contains(search)) ||
                    u.Name.Contains(search));

            var existingEmails = await _context.StudentWhiteLists
                .Select(s => s.Email.ToLower())
                .ToListAsync();

            var members = await query
                .OrderBy(u => u.Email)
                .Select(u => new
                {
                    u.Id,
                    u.Name,
                    Email = u.Email ?? string.Empty,
                })
                .ToListAsync();

            var result = members.Select(m => new
            {
                m.Id,
                m.Name,
                m.Email,
                alreadyAdded = existingEmails.Contains(m.Email.ToLower())
            });

            return Ok(result);
        }

        // POST /api/Student/whitelist - add a member to whitelist by userId (admin only)
        [HttpPost("whitelist")]
        public async Task<IActionResult> AddToWhiteList([FromBody] AddStudentRequest req)
        {
            if (!IsAdmin()) return Forbid();

            var adminEmail = User.FindFirstValue(ClaimTypes.Email) ?? string.Empty;

            var user = await _userManager.FindByIdAsync(req.UserId);
            if (user == null || string.IsNullOrEmpty(user.Email))
                return NotFound(new { error = "找不到此會員" });

            var exists = await _context.StudentWhiteLists
                .AnyAsync(s => s.Email.ToLower() == user.Email.ToLower());
            if (exists)
                return BadRequest(new { error = "此 email 已在白名單中" });

            var entry = new StudentWhiteList
            {
                Email = user.Email,
                Note = req.Note,
                AddedByEmail = adminEmail,
                AddedAt = DateTime.UtcNow
            };
            _context.StudentWhiteLists.Add(entry);
            await _context.SaveChangesAsync();

            return Ok(new { id = entry.Id, email = entry.Email, note = entry.Note });
        }

        // DELETE /api/Student/whitelist/{id} - remove from whitelist (admin only)
        [HttpDelete("whitelist/{id:int}")]
        public async Task<IActionResult> RemoveFromWhiteList(int id)
        {
            if (!IsAdmin()) return Forbid();

            var entry = await _context.StudentWhiteLists.FindAsync(id);
            if (entry == null) return NotFound();

            _context.StudentWhiteLists.Remove(entry);
            await _context.SaveChangesAsync();

            return Ok(new { deleted = true });
        }
    }
}
