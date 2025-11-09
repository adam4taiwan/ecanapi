// 檔案：Models/Analysis/EarthlyBranchHiddenStem.cs
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models.Analysis
{
    [Table("地支藏干")]
    public class EarthlyBranchHiddenStem
    {
        [Key]
        [Column("unique_id")]
        public int UniqueId { get; set; }

        [Column("floor")]
        public string? Floor { get; set; }

        [Column("sky1")]
        public string? Sky1 { get; set; }

        [Column("sky2")]
        public string? Sky2 { get; set; }

        [Column("sky3")]
        public string? Sky3 { get; set; }
    }
}