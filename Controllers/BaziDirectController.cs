using Ecanapi.Data;
using Ecanapi.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace Ecanapi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class BaziDirectController : ControllerBase
    {
        private readonly ApplicationDbContext _db;

        public BaziDirectController(ApplicationDbContext db)
        {
            _db = db;
        }

        // GET /api/BaziDirect/rules?ruleType=HourBranch&condition=子午卯酉
        [HttpGet("rules")]
        public async Task<IActionResult> GetRules(
            [FromQuery] string? ruleType = null,
            [FromQuery] int? chapter = null,
            [FromQuery] string? condition = null)
        {
            var query = _db.BaziDirectRules.AsQueryable();

            if (!string.IsNullOrEmpty(ruleType))
                query = query.Where(r => r.RuleType == ruleType);

            if (chapter.HasValue)
                query = query.Where(r => r.Chapter == chapter.Value);

            if (!string.IsNullOrEmpty(condition))
                query = query.Where(r => r.Condition.Contains(condition));

            var rules = await query
                .OrderBy(r => r.Chapter)
                .ThenBy(r => r.Section)
                .ThenBy(r => r.SortOrder)
                .ToListAsync();

            return Ok(rules);
        }

        // GET /api/BaziDirect/jianghu
        [HttpGet("jianghu")]
        public async Task<IActionResult> GetJianghuSecrets()
        {
            var rules = await _db.BaziDirectRules
                .Where(r => r.RuleType == "JianghuSecret")
                .OrderBy(r => r.SortOrder)
                .ToListAsync();

            return Ok(rules);
        }

        // GET /api/BaziDirect/hour-branch?branch=子午卯酉
        [HttpGet("hour-branch")]
        public async Task<IActionResult> GetHourBranchRule([FromQuery] string branch)
        {
            // Map the branch to one of the three groups
            string group = MapBranchToGroup(branch);
            if (string.IsNullOrEmpty(group))
                return BadRequest("Invalid branch");

            var rule = await _db.BaziDirectRules
                .Where(r => r.RuleType == "HourBranch" && r.Condition == group)
                .FirstOrDefaultAsync();

            return rule != null ? Ok(rule) : NotFound();
        }

        // GET /api/BaziDirect/hour-stem?stem=甲
        [HttpGet("hour-stem")]
        public async Task<IActionResult> GetHourStemRule([FromQuery] string stem)
        {
            string group = MapStemToGroup(stem);
            if (string.IsNullOrEmpty(group))
                return BadRequest("Invalid stem");

            var rule = await _db.BaziDirectRules
                .Where(r => r.RuleType == "HourStem" && r.Condition == group)
                .FirstOrDefaultAsync();

            return rule != null ? Ok(rule) : NotFound();
        }

        // GET /api/BaziDirect/source?chapter=1
        [HttpGet("source")]
        public async Task<IActionResult> GetSourceText([FromQuery] int? chapter = null)
        {
            var query = _db.FortuneSourceTexts.AsQueryable();
            if (chapter.HasValue)
                query = query.Where(t => t.ChapterNo == chapter.Value);

            var texts = await query.OrderBy(t => t.ChapterNo).ToListAsync();
            return Ok(texts);
        }

        // POST /api/BaziDirect/source (admin only)
        [HttpPost("source")]
        [Authorize]
        public async Task<IActionResult> AddSourceText([FromBody] FortuneSourceText text)
        {
            var userId = User.FindFirst(System.Security.Claims.ClaimTypes.NameIdentifier)?.Value;
            var user = await _db.Users.FindAsync(userId);
            if (user == null || user.Email != "adam4taiwan@gmail.com")
                return Forbid();

            text.Id = 0;
            text.CreatedAt = DateTime.UtcNow;
            _db.FortuneSourceTexts.Add(text);
            await _db.SaveChangesAsync();
            return Ok(text);
        }

        private static string MapBranchToGroup(string branch)
        {
            var group1 = new HashSet<string> { "子", "午", "卯", "酉" };
            var group2 = new HashSet<string> { "寅", "申", "巳", "亥" };
            var group3 = new HashSet<string> { "辰", "戌", "丑", "未" };

            if (group1.Contains(branch)) return "子午卯酉";
            if (group2.Contains(branch)) return "寅申巳亥";
            if (group3.Contains(branch)) return "辰戌丑未";
            return "";
        }

        private static string MapStemToGroup(string stem)
        {
            return stem switch
            {
                "甲" or "乙" => "甲乙時干",
                "丙" or "丁" => "丙丁時干",
                "戊" or "己" => "戊己時干",
                "庚" or "辛" => "庚辛時干",
                "壬" or "癸" => "壬癸時干",
                _ => ""
            };
        }
    }
}
