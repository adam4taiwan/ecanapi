using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models.Analysis
{
    [Table("starstyle")]
    public class StarStyle
    {
        [Key]
        [Column("unique_id")]
        public int UniqueId { get; set; }

        [Column("mapstar")]
        public string? MapStar { get; set; }

        [Column("mainstar")]
        public string? MainStar { get; set; }

        [Column("position")]
        public float? Position { get; set; }

        [Column("gd")]
        public string? Gd { get; set; }

        [Column("bd")]
        public string? Bd { get; set; }

        [Column("stardesc")]
        public string? StarDesc { get; set; }

        [Column("starbyyear")]
        public string? StarByYear { get; set; }
    }
}