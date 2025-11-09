using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models.Analysis
{
    [Table("六十甲子日對時")]
    public class SixtyJiaziDayToHour
    {
        [Key]
        [Column("unique_id")]
        public int UniqueId { get; set; }

        [Column("Sky")]
        public string? Sky { get; set; }

        [Column("Month")]
        public string? Month { get; set; }

        [Column("time")]
        public string? Time { get; set; }

        [Column("desc")]
        public string? Description { get; set; }
    }
}
