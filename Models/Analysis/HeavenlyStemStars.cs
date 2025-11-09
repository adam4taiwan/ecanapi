using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models.Analysis
{
    [Table("天干星剎")]
    public class HeavenlyStemStars
    {
        [Key]
        [Column("unique_id")]
        public int UniqueId { get; set; }

        [Column("KIND")]
        public string? Kind { get; set; }

        [Column("SKYNO")]
        public string? SkyNo { get; set; }

        [Column("TOFLO")]
        public string? ToFlo { get; set; }

        [Column("STAR")]
        public string? Star { get; set; }

        [Column("DESC")]
        public string? Description { get; set; }
    }
}
