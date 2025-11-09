// Controllers/ChartController.cs

using Microsoft.AspNetCore.Mvc;
using Ecanapi.Services; // 我們稍後會建立這個 using

namespace Ecanapi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ChartController : ControllerBase
    {
        private readonly IChartService _chartService;

        public ChartController(IChartService chartService)
        {
            _chartService = chartService;
        }

        [HttpPost("upload")]
        [ProducesResponseType(StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status400BadRequest)]
        [ProducesResponseType(StatusCodes.Status500InternalServerError)]
        public async Task<IActionResult> UploadChart(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest(new { message = "沒有上傳任何檔案。" });
            }

            var fileExtension = Path.GetExtension(file.FileName).ToLower();
            if (fileExtension != ".xls" && fileExtension != ".xlsx")
            {
                return BadRequest(new { message = "檔案格式不正確，請上傳 .xls 或 .xlsx 檔案。" });
            }

            try
            {
                var chartData = await _chartService.ProcessChartFileAsync(file);
                return Ok(chartData);
            }
            catch (Exception ex)
            {
                // 實際應用中應該記錄更詳細的錯誤日誌
                return StatusCode(500, new { message = $"內部伺-服器錯誤: {ex.Message}" });
            }
        }
    }
}