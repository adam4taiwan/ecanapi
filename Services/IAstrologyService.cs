using Ecanapi.Models;
using Ecanapi.Services.AstrologyEngine;

namespace Ecanapi.Services
{
    public interface IAstrologyService
    {
        Task<AstrologyChartResult> CalculateChartAsync(AstrologyRequest request);
        // interface & service 都改這個簽名
        List<AnnualLuck> GenerateAnnualLucks(int fromYear, int toYear, string dayStem);

        void ExportAnnualLucksJson(List<AnnualLuck> list, string filename);

        Task<string> GetAiAnalysisAsync(string chartDataJson, string analysisType, string userQuestion = "");
    }
}
