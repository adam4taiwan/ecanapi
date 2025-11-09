using Ecanapi.Models;

namespace Ecanapi.Services
{
    public interface IAstrologyService
    {
        Task<AstrologyChartResult> CalculateChartAsync(AstrologyRequest request);
    }
}
