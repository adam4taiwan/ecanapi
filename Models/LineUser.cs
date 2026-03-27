using System.ComponentModel.DataAnnotations;

namespace Ecanapi.Models
{
    /// <summary>LINE Bot 用戶生辰資料（與 Ecanapi 用戶系統獨立）</summary>
    public class LineUser
    {
        public int Id { get; set; }

        /// <summary>LINE userId（來自 Webhook event.source.userId）</summary>
        [Required]
        public string LineUserId { get; set; } = string.Empty;

        public int BirthYear { get; set; }
        public int BirthMonth { get; set; }
        public int BirthDay { get; set; }

        /// <summary>M = 男，F = 女</summary>
        [Required]
        public string Gender { get; set; } = string.Empty;

        /// <summary>本命星（計算後存入，避免重複計算）</summary>
        public int NatalStar { get; set; }

        /// <summary>LINE 顯示名稱（選填）</summary>
        public string? DisplayName { get; set; }

        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
    }
}
