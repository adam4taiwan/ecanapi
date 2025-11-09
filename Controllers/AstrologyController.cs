using Ecanapi.Models;
using Ecanapi.Services;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using System.IO;
using System.Threading.Tasks;

namespace Ecanapi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class AstrologyController : ControllerBase
    {
        private readonly IAstrologyService _astrologyService;
        private readonly IExcelExportService _excelExportService;
        private readonly IAnalysisReportService _analysisReportService; // <--- 【關鍵修正】補上這一行宣告
        private readonly IWebHostEnvironment _env;

        public AstrologyController(
            IAstrologyService astrologyService,
            IExcelExportService excelExportService,
            IAnalysisReportService analysisReportService, // <--- 注入新的服務
            IWebHostEnvironment env)
        {
            _astrologyService = astrologyService;
            _excelExportService = excelExportService;
            _analysisReportService = analysisReportService; // <--- 初始化
            _env = env;
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
        // 【新增】產生 DOCX 分析報告的 API 端點
        [HttpPost("analyze")]
        public async Task<IActionResult> AnalyzeChart([FromBody] AstrologyRequest request)
        {
            // 1. 呼叫現有的服務，產生一份完整的命盤資料
            var chartData = await _astrologyService.CalculateChartAsync(request);

            // 2. 將「命盤資料」和「原始請求」一併傳遞給分析服務
            var fileBytes = await _analysisReportService.GenerateReportAsync(chartData, request); // <--- 【修改】此處傳入 request
            
            // 3. 設定下載檔案的名稱
            var exportFileName = $"{chartData.UserName}_AnalysisReport.docx";

            // 4. 將 byte 陣列作為 Word 檔案回傳給前端，觸發下載
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", exportFileName);
        }
    }
}