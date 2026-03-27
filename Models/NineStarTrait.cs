using System.ComponentModel.DataAnnotations;

namespace Ecanapi.Models
{
    /// <summary>九星本命星特質 KB（9 筆，Gemini 自動補充）</summary>
    public class NineStarTrait
    {
        public int Id { get; set; }

        /// <summary>星號 1-9</summary>
        [Required]
        public int StarNumber { get; set; }

        /// <summary>星名，例：一白水星</summary>
        [Required]
        public string StarName { get; set; } = string.Empty;

        /// <summary>個性特質（KB 文字，空則由 Gemini 生成後回填）</summary>
        public string? Personality { get; set; }

        /// <summary>事業財運</summary>
        public string? Career { get; set; }

        /// <summary>感情人際</summary>
        public string? Relationship { get; set; }

        /// <summary>健康養生</summary>
        public string? Health { get; set; }

        /// <summary>吉方位</summary>
        public string LuckyDirection { get; set; } = string.Empty;

        /// <summary>吉顏色</summary>
        public string LuckyColor { get; set; } = string.Empty;

        /// <summary>幸運數字</summary>
        public int LuckyNumber { get; set; }

        public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
    }
}
