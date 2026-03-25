using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace Ecanapi.Models
{
    public enum DateType
    {
        solar,  // 陽曆
        lunar   // 農曆
    }

    // 陰陽曆互換請求
    public class DateExchangeRequest
    {
        [Description("陽曆 = solar，農曆 = lunar")]
        public DateType DateType { get; set; } = DateType.solar;

        [Description("年")]
        public int Year { get; set; }

        [Description("月")]
        public int Month { get; set; }

        [Description("日")]
        public int Day { get; set; }

        [Description("是否為農曆閏月（僅 lunar 模式有效，一般留 false）")]
        [DefaultValue(false)]
        public bool IsLeapMonth { get; set; } = false;
    }

    // API 請求的輸入模型
    public record AstrologyRequest(
        int Year,
        int Month,
        int Day,
        int Hour,
        int Minute,
        int Gender, // 1: 男, 2: 女
        string Name,
        string? DateType = "solar"  // solar: 陽曆, lunar: 農曆
    );

    // 單一柱位的詳細資訊
    public record PillarInfo(
        string HeavenlyStem,
        string EarthlyBranch,
        string HeavenlyStemLiuShen,
        string NaYin,
        List<string> HiddenStemLiuShen
    );

    // 八字資訊
    public record BaziInfo(
        PillarInfo YearPillar,
        PillarInfo MonthPillar,
        PillarInfo DayPillar,
        PillarInfo TimePillar,
        string AnimalSign,
        string DayMaster
    );

    // 【新增】用來儲存單一柱八字大運的資訊
    public record BaziLuckCycle(
        int StartAge,          // 起始歲數
        int EndAge,            // 結束歲數
        string HeavenlyStem,   // 大運天干
        string EarthlyBranch,  // 大運地支
        string LiuShen         // 大運十神
    );

//    public record BaziShensha(
//    int StartAge,          // 起始歲數
//    int EndAge,            // 結束歲數
//    string HeavenlyStem,   // 大運天干
//    string EarthlyBranch,  // 大運地支
//    string Stars         // 大運十神
//);

    // 紫微斗數宮位資訊
    public record ZiWeiPalace(
        int Index,
        string PalaceName,
        string PalaceStem,
        string EarthlyBranch,
        List<string> MajorStars,
        List<string> SecondaryStars,
        List<string> AnnualStarTransformations,
        string DecadeAgeRange,
        string LifeCycleStage,
        string MainStarBrightness,
        string PalaceStemTransformations,
        List<string> GoodStars,
        List<string> BadStars,
        List<string> SmallStars
    );

    // 易經卦象資訊
    public record IChingHexagram(
        string Name,
        string Description,
        List<string> YaoLines,
        string ChangedToHexagramName
    );

    // 【更新】完整的命盤結果，加入了 BaziLuckCycles 列表
    // 檔案: AstrologyModels.cs (請替換 AstrologyChartResult 的定義)

    // 【更新】完整的命盤結果，加入了 BaziLuckCycles 列表 和 BaziShensha 列表
    public record AstrologyChartResult(
        BaziInfo Bazi,
        List<ZiWeiPalace> palaces,
        string WuXingJuText,
        string MingZhu,
        string ShenZhu,
        List<BaziLuckCycle> BaziLuckCycles,
        // 【⭐ 必須新增此行，這是編譯器抱怨 Model 沒有 BaziShensha 的原因 ⭐】
        List<string>? BaziShensha,
        IChingHexagram? InnateHexagram,
        IChingHexagram? PostnatalHexagram,
        string UserName,
        DateTime SolarBirthDate,
        string LunarBirthDate,
        ShiShenResult? BaziAnalysisResult,
        string? LuckCycleNote // 大運起運節氣計算說明
    );
    public class AnnualLuck
    {
        public int Year { get; set; }                // 流年西元
        public string HeavenlyStem { get; set; }     // 天干
        public string EarthlyBranch { get; set; }    // 地支
        public string StemLiuShen { get; set; }      // 天干六神
        public List<string>? BranchLiuShen { get; set; }   // 地支六神
        public List<string> YearShensha { get; set; } // 流年神煞
        public List<string> BranchRelations { get; set; } // 刑沖害合說明
        public List<string> AffectBazi { get; set; }      // 影響八字本命
        public string MajorLuckStem { get; set; }    // 大運天干
        public string MajorLuckBranch { get; set; }  // 大運地支
        public List<string>? AnnualInteractions { get; set; } = new List<string>(); // 【新增 2：流年與命盤的刑沖合害】
    }

}