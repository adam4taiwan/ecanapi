using System.ComponentModel.DataAnnotations;

namespace Ecanapi.Models
{
    public class UserChart
    {
        public int Id { get; set; }

        [Required]
        public string UserId { get; set; } = string.Empty;

        /// <summary>命宮主星（逗號分隔，例如 "紫微,天府"）</summary>
        public string? MingGongMainStars { get; set; }

        /// <summary>用神五行（木火土金水），由 analyze 端點計算後回填，供每日推播喜忌判定</summary>
        [MaxLength(2)]
        public string? YongShenElem { get; set; }

        /// <summary>忌神五行</summary>
        [MaxLength(2)]
        public string? JiShenElem { get; set; }

        /// <summary>格局名稱（建祿格/月刃格/食神格等）</summary>
        [MaxLength(20)]
        public string? Pattern { get; set; }

        /// <summary>日主身強弱百分比（0-100），由 analyze 端點回填</summary>
        public int? BodyPct { get; set; }

        /// <summary>完整命盤 JSON（AstrologyResult）</summary>
        [Required]
        public string ChartJson { get; set; } = string.Empty;

        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
    }
}
