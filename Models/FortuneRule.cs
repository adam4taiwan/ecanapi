using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models
{
    [Table("FortuneRules")]
    public class FortuneRule
    {
        public int Id { get; set; }

        [Required]
        public string Category { get; set; } = string.Empty;   // 八字 / 紫微 / 通用

        public string? Subcategory { get; set; }                // 四化 / 格局 / 星情 / 宮位 / 大運 etc.

        public string? Title { get; set; }                      // 條目標題

        public string? ConditionText { get; set; }              // 觸發條件描述

        [Required]
        public string ResultText { get; set; } = string.Empty;  // 論斷內容

        public string? SourceFile { get; set; }                 // 來源文件

        public string? Tags { get; set; }                       // 以逗號分隔的標籤

        public bool IsActive { get; set; } = true;

        public int SortOrder { get; set; } = 0;

        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

        public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
    }

    [Table("KnowledgeDocuments")]
    public class KnowledgeDocument
    {
        public int Id { get; set; }

        [Required]
        public string FileName { get; set; } = string.Empty;

        [Required]
        public string FileType { get; set; } = string.Empty;   // csv / docx / doc / xlsx / txt / pdf

        public string? Category { get; set; }

        public string? ContentPreview { get; set; }             // 前500字預覽

        public int RuleCount { get; set; } = 0;

        public string Status { get; set; } = "imported";        // imported / error

        public DateTime UploadedAt { get; set; } = DateTime.UtcNow;

        public string? UploadedBy { get; set; }
    }
}
