// 檔案：Models/Analysis/BodyMaster.cs
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models.Analysis
{
    [Table("身主")]
    public class BodyMaster
    {
        [Key]
        [Column("unique_id")]
        public int UniqueId { get; set; }

        [Column("floor")]
        public string? Floor { get; set; }

        [Column("star")]
        public string? Star { get; set; }
    }
}