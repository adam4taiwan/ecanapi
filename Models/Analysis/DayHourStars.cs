using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models.Analysis
{
    [Table("日對時星剎")]
    public class DayHourStars
    {
        [Key]
        [Column("unique_id")]
        public int UniqueId { get; set; }

        [Column("SkyFloor")]
        public string? SkyFloor { get; set; }

        [Column("position")]
        public string? Position { get; set; }

        [Column("desc")]
        public string? Description { get; set; }
    }
}
