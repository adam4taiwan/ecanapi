using System.ComponentModel.DataAnnotations;

namespace Ecanapi.Models
{
    public class DailyFortune
    {
        public int Id { get; set; }

        [Required]
        public string UserId { get; set; } = string.Empty;

        /// <summary>運勢日期（只用日期部分）</summary>
        public DateTime FortuneDate { get; set; }

        /// <summary>AI 生成的運勢內容</summary>
        [Required]
        public string Content { get; set; } = string.Empty;

        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }
}
