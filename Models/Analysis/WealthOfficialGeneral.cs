// 檔案：Models/Analysis/WealthOfficialGeneral.cs
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models.Analysis
{
    [Table("財官總論")]
    public class WealthOfficialGeneral
    {
        [Key]
        [Column("unique_id")]
        public int UniqueId { get; set; }

        [Column("sky")]
        public string? Sky { get; set; }

        [Column("desc")]
        public string? Description { get; set; }
    }
}