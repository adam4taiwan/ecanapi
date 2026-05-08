using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models
{
    [Table("六神四柱口訣")]
    public class BaziPillarFormula
    {
        [Key]
        public int Id { get; set; }

        // e.g. "月干財", "年干才", "日干殺"
        public string? Position { get; set; }

        // full ten god name, e.g. "正財", "偏財", "偏官"
        public string? Star { get; set; }

        // short code: 財/才/官/殺/印/ㄗ/食/傷/比/劫
        public string? Simple { get; set; }

        // pillar: 年/月/日/時
        public string? Pillar { get; set; }

        // main description text
        public string? Gd { get; set; }

        // newer/refined description (preferred)
        [Column("newdesc")]
        public string? NewDesc { get; set; }

        public int Uid { get; set; }
    }
}
