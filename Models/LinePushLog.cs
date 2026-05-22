using System.ComponentModel.DataAnnotations;

namespace Ecanapi.Models
{
    /// <summary>LINE Bot 每日推播記錄，供 AI 讀取作為情境參考</summary>
    public class LinePushLog
    {
        public int Id { get; set; }

        /// <summary>LINE userId</summary>
        [Required]
        [MaxLength(100)]
        public string LineUserId { get; set; } = string.Empty;

        /// <summary>對應 AspNetUsers.Id（訂閱會員才有）</summary>
        [MaxLength(450)]
        public string? UserId { get; set; }

        /// <summary>用戶 email（便於查詢，不強制唯一）</summary>
        [MaxLength(256)]
        public string? UserEmail { get; set; }

        /// <summary>推播類型：ninestar（純九星用戶）/ subscriber（訂閱會員個人化）</summary>
        [MaxLength(20)]
        public string PushType { get; set; } = string.Empty;

        /// <summary>本命星 1-9</summary>
        public int NatalStar { get; set; }

        /// <summary>出生年（AI 情境用）</summary>
        public int? BirthYear { get; set; }

        /// <summary>日主天干（AI 個人化情境用）</summary>
        [MaxLength(2)]
        public string? DayMaster { get; set; }

        /// <summary>推播日期（台灣時間，YYYY-MM-DD）</summary>
        public DateOnly PushDate { get; set; }

        /// <summary>推播完整文字內容</summary>
        [Required]
        public string Message { get; set; } = string.Empty;

        /// <summary>實際發送時間（UTC）</summary>
        public DateTime SentAt { get; set; } = DateTime.UtcNow;

        /// <summary>發送結果：success / failed</summary>
        [MaxLength(10)]
        public string Status { get; set; } = string.Empty;

        /// <summary>失敗原因（失敗時記錄）</summary>
        public string? ErrorMessage { get; set; }
    }
}
