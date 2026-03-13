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



    }
}