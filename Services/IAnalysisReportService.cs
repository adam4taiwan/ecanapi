using Ecanapi.Models;
using System.Threading.Tasks;

namespace Ecanapi.Services
{
    /// <summary>
    /// 定義命盤分析與報告產出服務的標準。
    /// </summary>
    public interface IAnalysisReportService
    {
        /// <summary>
        /// 根據命盤資料，生成一份完整的 DOCX 分析報告。
        /// </summary>
        /// <param name="chartData">已排好的命盤資料。</param>
        /// <param name="request">原始的 API 請求資料。</param> // <--- 【新增】參數
        /// <returns>包含 DOCX 檔案內容的位元組陣列。</returns>
        Task<byte[]> GenerateReportAsync(AstrologyChartResult chartData, AstrologyRequest request); // <--- 【修改】方法簽名
    }
}