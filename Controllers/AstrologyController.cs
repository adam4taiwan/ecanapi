using Ecanapi.Data;
using Ecanapi.Models;
using Ecanapi.Services;
using Ecanapi.Services.AstrologyEngine;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.IO;
using System.Security.Claims;
using System.Text.Json;
using System.Threading.Tasks;

namespace Ecanapi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class AstrologyController : ControllerBase
    {
        private readonly IAstrologyService _astrologyService;
        private readonly IExcelExportService _excelExportService;
        private readonly IAnalysisReportService _analysisReportService;
        private readonly IWebHostEnvironment _env;
        private readonly ApplicationDbContext _context;

        public AstrologyController(
            IAstrologyService astrologyService,
            IExcelExportService excelExportService,
            IAnalysisReportService analysisReportService,
            IWebHostEnvironment env,
            ApplicationDbContext context)
        {
            _astrologyService = astrologyService;
            _excelExportService = excelExportService;
            _analysisReportService = analysisReportService;
            _env = env;
            _context = context;
        }

        [HttpPost("calculate")]
        public async Task<IActionResult> CalculateChart([FromBody] AstrologyRequest request)
        {
            if (request == null)
            {
                return BadRequest("Invalid request data.");
            }
            var result = await _astrologyService.CalculateChartAsync(request);
            return Ok(result);
        }

        [HttpPost("calculateFull")]
        public async Task<IActionResult> CalculateFull([FromBody] AstrologyRequest request, int fromYear, int toYear)
        {
            // 呼叫原本的八字命盤分析
            var context = new AstrologyCalculationContext(request);
            var chartResult = await _astrologyService.CalculateChartAsync(request);
            BaziInfo bazi = chartResult.Bazi;
            string dayStem = bazi.DayPillar.HeavenlyStem;
            // 從八字 context 取得 cue3（如 context.CUE3 為日主天干index）
            //string[] GAN = { "甲", "乙", "丙", "丁", "戊", "己", "庚", "辛", "壬", "癸" };
            //string dayStem = GAN[context.CUE3];

            // 呼叫新批次流年六神
            var annualLuckList = _astrologyService.GenerateAnnualLucks(fromYear, toYear, dayStem);

            // 回傳合併內容
            return Ok(new
            {
                Chart = chartResult,
                AnnualLucks = annualLuckList
            });
        }

        [HttpPost("export")]
        public async Task<IActionResult> ExportChart([FromBody] AstrologyRequest request)
        {
            if (request == null)
            {
                return BadRequest("Invalid request data.");
            }

            // 正常流程：呼叫修正後的 AstrologyService 產生正確資料
            var chartData = await _astrologyService.CalculateChartAsync(request);

            var templatePath = Path.Combine(_env.WebRootPath, "templates", "chart_template.xlt");

            var fileBytes = await _excelExportService.GenerateChartAsync(chartData, templatePath);

            var exportFileName = $"{chartData.UserName}_AstrologyChart.xls";

            return File(fileBytes, "application/vnd.ms-excel", exportFileName);
        }
        // 產生專家命書 DOCX 報告（需登入，扣 200 點）
        [HttpPost("analyze")]
        [Authorize]
        public async Task<IActionResult> AnalyzeChart([FromBody] AstrologyRequest request)
        {
            const int cost = 200;
            var identity = User.FindFirstValue(ClaimTypes.Email) ?? User.FindFirstValue(ClaimTypes.Name);
            if (string.IsNullOrEmpty(identity))
                return Unauthorized(new { error = "請重新登入" });

            var user = await _context.Users.FirstOrDefaultAsync(u => u.Email == identity || u.UserName == identity);
            if (user == null) return Unauthorized(new { error = "找不到用戶" });

            var adminEmail = HttpContext.RequestServices.GetRequiredService<IConfiguration>()["Admin:Email"];
            if (!string.Equals(user.Email, adminEmail, StringComparison.OrdinalIgnoreCase))
                return StatusCode(403, new { error = "此功能僅限管理員使用" });

            if (user.Points < cost) return BadRequest(new { error = $"點數不足，此功能需要 {cost} 點" });

            var chartData = await _astrologyService.CalculateChartAsync(request);
            var fileBytes = await _analysisReportService.GenerateReportAsync(chartData, request);

            user.Points -= cost;
            await _context.SaveChangesAsync();

            var exportFileName = $"{chartData.UserName}_ExpertReport.docx";
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", exportFileName);
        }
        [HttpPost("ExchangeDate")]
        public IActionResult ExchangeDate([FromBody] DateExchangeRequest request)
        {
            if (request == null)
                return BadRequest("Invalid request data.");

            try
            {
                if (request.DateType == DateType.solar)
                {
                    var solarDate = new DateTime(request.Year, request.Month, request.Day);
                    var cal = new Ecan.EcanChineseCalendar(solarDate);
                    return Ok(new
                    {
                        inputType = "solar (陽曆)",
                        input = new { year = request.Year, month = request.Month, day = request.Day },
                        output = new
                        {
                            year = cal.ChineseYear,
                            month = cal.ChineseMonth,
                            day = cal.ChineseDay,
                            isLeapMonth = cal.IsChineseLeapMonth,
                            yearString = cal.ChineseYearString,
                            monthString = cal.ChineseMonthString,
                            dayString = cal.ChineseDayString
                        }
                    });
                }
                else if (request.DateType == DateType.lunar)
                {
                    var cal = new Ecan.EcanChineseCalendar(request.Year, request.Month, request.Day, request.IsLeapMonth);
                    var solar = cal.Date;
                    return Ok(new
                    {
                        inputType = "lunar (農曆)",
                        input = new { year = request.Year, month = request.Month, day = request.Day, isLeapMonth = request.IsLeapMonth },
                        output = new
                        {
                            year = solar.Year,
                            month = solar.Month,
                            day = solar.Day
                        }
                    });
                }
                else
                {
                    return BadRequest(new { error = "dateType must be 'solar' or 'lunar'" });
                }
            }
            catch (Exception ex)
            {
                return BadRequest(new { error = ex.Message });
            }
        }

        [HttpGet("ExportAnnualLuckJson")]
        public IActionResult ExportAnnualLuckJson(int fromYear, int toYear, string fileName, int cue3)
        {
            string[] GAN = { "甲", "乙", "丙", "丁", "戊", "己", "庚", "辛", "壬", "癸" };
            string dayStem = GAN[cue3]; // cue3 為日干從 0~9
            var data = _astrologyService.GenerateAnnualLucks(fromYear, toYear, dayStem);
            string savePath = Path.Combine("D:\\AstroOutput", fileName + ".json");
            _astrologyService.ExportAnnualLucksJson(data, savePath);
            return Ok(new { output = savePath, count = data.Count });
        }

        /// <summary>儲存用戶命盤（供 Phase 3 紫微命宮加成使用）</summary>
        [HttpPost("save-chart")]
        [Authorize]
        public async Task<IActionResult> SaveChart([FromBody] JsonElement chartJson)
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (string.IsNullOrEmpty(userId))
                return Unauthorized(new { error = "請重新登入" });

            // 從 chartJson 中提取命宮主星
            string mingGongMainStars = ExtractMingGongMainStars(chartJson);

            string jsonStr = chartJson.GetRawText();

            // 更新或新增
            var existing = await _context.UserCharts.FirstOrDefaultAsync(c => c.UserId == userId);
            if (existing != null)
            {
                existing.MingGongMainStars = mingGongMainStars;
                existing.ChartJson = jsonStr;
                existing.UpdatedAt = DateTime.UtcNow;
            }
            else
            {
                _context.UserCharts.Add(new UserChart
                {
                    UserId = userId,
                    MingGongMainStars = mingGongMainStars,
                    ChartJson = jsonStr,
                    CreatedAt = DateTime.UtcNow,
                    UpdatedAt = DateTime.UtcNow
                });
            }

            await _context.SaveChangesAsync();
            return Ok(new { success = true, mingGongMainStars });
        }

        /// <summary>取得用戶已儲存的命盤</summary>
        [HttpGet("my-chart")]
        [Authorize]
        public async Task<IActionResult> GetMyChart()
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (string.IsNullOrEmpty(userId))
                return Unauthorized(new { error = "請重新登入" });

            var chart = await _context.UserCharts.FirstOrDefaultAsync(c => c.UserId == userId);
            if (chart == null)
                return NotFound(new { error = "尚未儲存命盤" });

            return Ok(new
            {
                mingGongMainStars = chart.MingGongMainStars,
                updatedAt = chart.UpdatedAt
            });
        }

        private static readonly Dictionary<string, string> StarAbbrevMap = new()
        {
            {"紫","紫微"},{"機","天機"},{"陽","太陽"},{"武","武曲"},
            {"同","天同"},{"廉","廉貞"},{"府","天府"},{"陰","太陰"},
            {"貪","貪狼"},{"巨","巨門"},{"相","天相"},{"梁","天梁"},
            {"殺","七殺"},{"破","破軍"}
        };

        private static string ExtractMingGongMainStars(JsonElement chartJson)
        {
            try
            {
                // chartJson 結構：{ palaces: [ { palaceName, majorStars, ... } ] }
                JsonElement palacesEl;

                if (chartJson.TryGetProperty("ziwei", out var ziweiEl) &&
                    ziweiEl.TryGetProperty("palaces", out palacesEl))
                {
                    // nested under ziwei
                }
                else if (chartJson.TryGetProperty("Ziwei", out var ziweiElCap) &&
                         ziweiElCap.TryGetProperty("Palaces", out palacesEl))
                {
                    // PascalCase
                }
                else if (chartJson.TryGetProperty("palaces", out palacesEl) ||
                         chartJson.TryGetProperty("Palaces", out palacesEl))
                {
                    // top-level palaces
                }
                else
                {
                    return string.Empty;
                }

                foreach (var palace in palacesEl.EnumerateArray())
                {
                    string palaceName = string.Empty;
                    if (palace.TryGetProperty("palaceName", out var pn)) palaceName = pn.GetString() ?? "";
                    else if (palace.TryGetProperty("PalaceName", out var pnCap)) palaceName = pnCap.GetString() ?? "";

                    if (palaceName != "命宮") continue;

                    JsonElement starsEl;
                    if (!palace.TryGetProperty("majorStars", out starsEl) &&
                        !palace.TryGetProperty("MajorStars", out starsEl))
                        return string.Empty;

                    var stars = new List<string>();
                    foreach (var s in starsEl.EnumerateArray())
                    {
                        var abbr = s.GetString();
                        if (string.IsNullOrEmpty(abbr)) continue;
                        // 嘗試展開縮寫為全名
                        stars.Add(StarAbbrevMap.TryGetValue(abbr, out var full) ? full : abbr);
                    }
                    return string.Join(",", stars);
                }
            }
            catch
            {
                // ignore parse errors
            }
            return string.Empty;
        }
    }
}