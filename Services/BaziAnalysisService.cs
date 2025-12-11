// Services/BaziAnalysisService.cs (請修改此檔案)

using Ecanapi.Models;
using System.Threading.Tasks;

// 【⭐⭐ 確保在正確的命名空間內 ⭐⭐】
namespace Ecanapi.Services
{
    // 繼承自介面 IBaziAnalysisService
    public class BaziAnalysisService : IBaziAnalysisService
    {
        // 【新增】將分析器宣告為私有唯讀欄位，確保它在服務中隨時可用
        private readonly BlindSchoolUltimateAnalyzer _analyzer = new();

        public async Task<BlindSchoolAnalysisResult> AnalyzeAsync(BaziInfo bazi)
        {
            // 1. 將 BaziInfo 中的四柱資訊轉換成干支字串 (Analyzer需要的輸入)
            string yearPillar = bazi.YearPillar.HeavenlyStem + bazi.YearPillar.EarthlyBranch;
            string monthPillar = bazi.MonthPillar.HeavenlyStem + bazi.MonthPillar.EarthlyBranch;
            string dayPillar = bazi.DayPillar.HeavenlyStem + bazi.DayPillar.EarthlyBranch;
            string timePillar = bazi.TimePillar.HeavenlyStem + bazi.TimePillar.EarthlyBranch;

            // 2. 呼叫核心分析器的同步方法
            var result = _analyzer.Analyze(yearPillar, monthPillar, dayPillar, timePillar);

            // 3. 回傳結果
            return result;
        }
    }
}