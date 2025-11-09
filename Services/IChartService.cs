// Services/IChartService.cs (已更新回傳型別)

using Ecanapi.Models; // 1. 引用我們新的 Model

namespace Ecanapi.Services
{
    public interface IChartService
    {
        // 2. 將回傳型別從 object 改為 ChartDataDto，與 ChartService.cs 保持一致
        Task<ChartDataDto> ProcessChartFileAsync(IFormFile file);
    }
}