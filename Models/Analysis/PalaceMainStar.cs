// 檔案：Models/Analysis/PalaceMainStar.cs
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models.Analysis
{
    [Table("命宮主星")]
    public class PalaceMainStar
    {
        [Key]
        [Column("unique_id")]
        public int UniqueId { get; set; }

        [Column("star")]
        public string? Star { get; set; }

        [Column("star_type")]
        public string? StarType { get; set; }

        [Column("star_desc")]
        public string? StarDesc { get; set; }

        [Column("star_value")]
        public string? StarValue { get; set; }
    }
}