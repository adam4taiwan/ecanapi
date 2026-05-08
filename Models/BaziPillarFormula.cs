using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models
{
    [Table("六神四柱口訣")]
    public class BaziPillarFormula
    {
        [Key]
        [Column("id")]
        public int Id { get; set; }

        [Column("position")]
        public string? Position { get; set; }

        [Column("star")]
        public string? Star { get; set; }

        [Column("simple")]
        public string? Simple { get; set; }

        [Column("pillar")]
        public string? Pillar { get; set; }

        [Column("gd")]
        public string? Gd { get; set; }

        [Column("newdesc")]
        public string? NewDesc { get; set; }

        [Column("uid")]
        public int Uid { get; set; }
    }
}
