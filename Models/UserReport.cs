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

        /// <summary>審核狀態: pending_review / approved / rejected</summary>
        [MaxLength(20)]
        public string Status { get; set; } = "pending_review";

        public DateTime? ApprovedAt { get; set; }

        /// <summary>管理員備註（退回原因等）</summary>
        [Column(TypeName = "text")]
        public string? AdminNote { get; set; }

        /// <summary>時效下載 token（核准送出時產生，72小時有效）</summary>
        [MaxLength(64)]
        public string? DownloadToken { get; set; }

        public DateTime? DownloadTokenExpiry { get; set; }

        /// <summary>管理員上傳的最終修正版 DOCX 二進位（核准後供用戶下載）</summary>
        [Column(TypeName = "bytea")]
        public byte[]? ApprovedDocxBytes { get; set; }

        /// <summary>最終 DOCX 原始檔名</summary>
        [MaxLength(200)]
        public string? ApprovedDocxFileName { get; set; }
    }
}
