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

        // ── 對話狀態機 ──────────────────────────────
        /// <summary>對話狀態：idle / reg_year / reg_month / reg_day / reg_gender</summary>
        [MaxLength(20)]
        public string State { get; set; } = "idle";

        /// <summary>設定生辰流程暫存：年</summary>
        public int? TempYear { get; set; }

        /// <summary>設定生辰流程暫存：月</summary>
        public int? TempMonth { get; set; }

        /// <summary>設定生辰流程暫存：日</summary>
        public int? TempDay { get; set; }

        // ── 推播訂閱 ────────────────────────────────
        /// <summary>是否訂閱每日七星開運通知</summary>
        public bool NotifyEnabled { get; set; } = false;

        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
    }
}
