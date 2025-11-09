// 檔案：Models/Analysis/PalaceName.cs
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models.Analysis
{
    [Table("十二宮稱呼")]
    public class PalaceName
    {
        [Key]
        [Column("unique_id")]
        public int UniqueId { get; set; }

        [Column("name")]
        public string? Name { get; set; }

        [Column("desc")]
        public string? Description { get; set; }
    }
}