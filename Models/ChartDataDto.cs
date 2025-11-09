// Models/ChartDataDto.cs (已擴充八字結構)

namespace Ecanapi.Models
{
    // 用於表示單一「柱」的詳細資訊
    public class BaziPillarDto
    {
        public string? HeavenlyStem { get; set; }         // 天干
        public string? EarthlyBranch { get; set; }        // 地支
        public string? HeavenlyStemLiuShen { get; set; }  // 天干六神
        public string? NaYin { get; set; }                  // 納音
        public List<string> HiddenStemLiuShen { get; set; } = new(); // 地支藏干六神
    }

    public class PalaceDto
    {
        public string? Name { get; set; }
        public string? EarthlyBranch { get; set; }
        public bool IsBodyPalace { get; set; } = false; // <--- 新增的身宮標記欄位
        public List<string> MainStars { get; set; } = new();
        public List<string> SecondaryStars { get; set; } = new();
        public List<string> MinorStars { get; set; } = new();
        public string? MainStarBrightness { get; set; }
        public string? PalaceAuspiciousness { get; set; }
        public List<string> AnnualStarTransformations { get; set; } = new();
        public string? PalaceStem { get; set; }
        public string? PalaceStemTransformations { get; set; }
        public List<string> YearlyGeneralStars { get; set; } = new();
        public string? LifeCycleStage { get; set; }
    }

    // 將原有的 BaziDto 升級為包含四個詳細的柱
    public class BaziDto
    {
        public BaziPillarDto YearPillar { get; set; } = new();
        public BaziPillarDto MonthPillar { get; set; } = new();
        public BaziPillarDto DayPillar { get; set; } = new();
        public BaziPillarDto TimePillar { get; set; } = new();
    }

    public class ChartDataDto
    {
        public string? FileName { get; set; }
        public string? Name { get; set; }
        public string? BirthDate { get; set; }
        public BaziDto BaziInfo { get; set; } = new();
        public List<PalaceDto> Palaces { get; set; } = new();
        public List<LuckCyclePeriodDto> DecadeLuckCycles { get; set; } = new();
    }

    public class LuckCyclePeriodDto
    {
        public int StartAge { get; set; }
        public int EndAge { get; set; }
        public string? AssociatedPalace { get; set; }
    }
}