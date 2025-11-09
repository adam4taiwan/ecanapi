// 檔案：Models/Analysis/PalaceStarBrightness.cs
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models.Analysis
{
    [Table("十二宮廟旺")]
    public class PalaceStarBrightness
    {
        [Key]
        [Column("unique_id")]
        public int UniqueId { get; set; }

        [Column("star")]
        public string? Star { get; set; }

        [Column("palace")]
        public string? Palace { get; set; }

        [Column("light")]
        public string? Light { get; set; }
    }
}