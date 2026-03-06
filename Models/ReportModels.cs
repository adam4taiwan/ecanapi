using System;
using System.Collections.Generic;

namespace Ecanapi.Models
{
    // API 請求的輸入模型
    public enum EnergyLevel
    {
        Long生 = 12, Mu = 3, Jue = 1, // 長生、墓、絕等
        Wang = 10, // 帝旺
        KongWang = 0 // 空亡修正
    }
    public class StarStyleDesc
    {
        public string MainStar { get; set; } = "";
        public string Position { get; set; } = "";
        public string Gd { get; set; } = "";
        public string Bd { get; set; } = "";
        public string StarDesc { get; set; } = "";
        public string StarByYear { get; set; } = "";
    }
    public class PillarAnalysisResult
    {
        public string DayPillar { get; set; }
        public List<string> Diagnoses { get; set; } = new List<string>();
        public string MarriageStatus { get; set; }
        public string ChildrenStatus { get; set; }
        // 補上這一個欄位，解決編譯錯誤
        public string CareerStatus { get; set; }
        // --- 補上這一個欄位，解決 CS1061 錯誤 ---
        /// <summary>
        /// 日元性情分析 (北派盲訣內容)
        /// </summary>
        public string DayMasterAnalysis { get; set; }
        public string RelativesAnalysis { get; set; }
    }
}