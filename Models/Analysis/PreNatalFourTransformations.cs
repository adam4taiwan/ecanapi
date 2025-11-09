using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models.Analysis
{
    [Table("先天四化入十二宮")]
    public class PreNatalFourTransformations
    {
        [Key]
        [Column("unique_id")]
        public int UniqueId { get; set; }

        [Column("mainstar")]
        public string? MainStar { get; set; }

        [Column("position")]
        public int? Position { get; set; }

        [Column("desc")]
        public string? Description { get; set; }
    }
}
