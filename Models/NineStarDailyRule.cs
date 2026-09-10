using System.ComponentModel.DataAnnotations;

namespace Ecanapi.Models
{
    /// <summary>九星每日建議 KB（本命星×流年×流月×流日，Gemini 自動補充）</summary>
    public class NineStarDailyRule
    {
        public int Id { get; set; }

        /// <summary>本命星 1-9</summary>
        [Required]
        public int NatalStar { get; set; }

        /// <summary>流年星 1-9（0=舊版紀錄，無年月脈絡）</summary>
        public int YearStar { get; set; } = 0;

        /// <summary>流月星 1-9（0=舊版紀錄，無年月脈絡）</summary>
        public int MonthStar { get; set; } = 0;

        /// <summary>流日星 1-9</summary>
        [Required]
        public int FlowStar { get; set; }

        /// <summary>運勢說明（空則由 Gemini 生成後回填）</summary>
        public string? FortuneText { get; set; }

        /// <summary>今日宜</summary>
        public string? Auspicious { get; set; }

        /// <summary>今日忌</summary>
        public string? Avoid { get; set; }

        /// <summary>吉方位</summary>
        public string? Direction { get; set; }

        /// <summary>吉顏色</summary>
        public string? Color { get; set; }

        public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
    }
}
