using System;
using System.Collections.Generic;

namespace Ecanapi.Models
{
    // API 請求的輸入模型
    public record AstrologyRequest(
        int Year,
        int Month,
        int Day,
        int Hour,
        int Minute,
        int Gender, // 1: 男, 2: 女
        string Name
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
    public record AstrologyChartResult(
        BaziInfo Bazi,
        List<ZiWeiPalace> palaces,
        string WuXingJuText,
        string MingZhu,
        string ShenZhu,
        List<BaziLuckCycle> BaziLuckCycles, // <--- 新增八字大運列表
        IChingHexagram? InnateHexagram,
        IChingHexagram? PostnatalHexagram,
        string UserName,
        DateTime SolarBirthDate,
        string LunarBirthDate
    );
}