using Ecanapi.Models;
using System.Threading.Tasks;

namespace Ecanapi.Services
{
    /// <summary>
    /// 定義命盤分析與報告產出服務的標準。
    /// </summary>
    public interface IProReportService
    {
        Task<byte[]> GenerateProReportAsync(AstrologyChartResult chartData, AstrologyRequest request);
    }
}