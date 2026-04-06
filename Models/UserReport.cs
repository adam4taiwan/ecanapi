using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models
{
    /// <summary>使用者命書記錄（每次生成命書時存檔）</summary>
    public class UserReport
    {
        [Key]
        public Guid Id { get; set; } = Guid.NewGuid();

        [Required]
        public required string UserId { get; set; }

        /// <summary>命書類型: bazi / daiyun / liunian / lifelong</summary>
        [MaxLength(20)]
        public required string ReportType { get; set; }

        /// <summary>顯示標題，例如「2026 流年命書」、「八字命書」</summary>
        [MaxLength(100)]
        public required string Title { get; set; }

        /// <summary>完整命書文字（Markdown 格式）</summary>
        [Column(TypeName = "text")]
        public required string Content { get; set; }

        /// <summary>生成參數 JSON（生辰、年份等，供未來重算參考）</summary>
        [Column(TypeName = "jsonb")]
        public string? Parameters { get; set; }

        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }
}
