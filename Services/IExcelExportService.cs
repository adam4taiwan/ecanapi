using Ecanapi.Models;
using System.Threading.Tasks;

namespace Ecanapi.Services
{
    public interface IExcelExportService
    {
        Task<byte[]> GenerateChartAsync(AstrologyChartResult chartData, string templatePath);
    }
}