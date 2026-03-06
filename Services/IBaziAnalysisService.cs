// Services/IBaziAnalysisService.cs (新增此檔案)
using Ecanapi.Models;
using System.Threading.Tasks;

namespace Ecanapi.Services
{
    public interface IBaziAnalysisService
    {
        // 接受 BaziInfo，回傳分析結果模型
        Task<BlindSchoolAnalysisResult> AnalyzeAsync(BaziInfo bazi);
    }
}
