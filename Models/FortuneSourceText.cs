using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models
{
    [Table("FortuneSourceTexts")]
    public class FortuneSourceText
    {
        [Key]
        public int Id { get; set; }

        [MaxLength(200)]
        public string Title { get; set; } = "";

        // 0 = intro, 1~12 = chapter number
        public int ChapterNo { get; set; }

        // Full text for reference lookup
        public string FullText { get; set; } = "";

        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }
}
