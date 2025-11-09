// 檔案：Models/Analysis/HeavenlyStemInfo.cs
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models.Analysis
{
    [Table("天干陰陽五行")]
    public class HeavenlyStemInfo
    {
        [Key]
        [Column("unique_id")]
        public int UniqueId { get; set; }

        [Column("sky")]
        public string? Sky { get; set; }

        [Column("yin_yang")]
        public string? YinYang { get; set; }

        [Column("element")]
        public string? Element { get; set; }
    }
}