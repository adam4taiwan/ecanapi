// 檔案：Models/Analysis/StarCondition.cs
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models.Analysis
{
    [Table("星曜狀況")]
    public class StarCondition
    {
        [Key]
        [Column("unique_id")]
        public int UniqueId { get; set; }

        [Column("star")]
        public string? Star { get; set; }

        [Column("desc")]
        public string? Description { get; set; }
    }
}