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

        /// <summary>完整命盤 JSON（AstrologyResult）</summary>
        [Required]
        public string ChartJson { get; set; } = string.Empty;

        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
    }
}
