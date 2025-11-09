// 檔案：Models/Analysis/IChingExplanation.cs
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models.Analysis
{
    [Table("易經六十四卦分類解說")]
    public class IChingExplanation
    {
        [Key]
        [Column("unique_id")]
        public int UniqueId { get; set; }

        [Column("gua_id")]
        public int? GuaId { get; set; }

        [Column("gua_value")]
        public int? GuaValue { get; set; }

        [Column("gua_desc")]
        public string? GuaDesc { get; set; }

        [Column("gua_type")]
        public string? GuaType { get; set; }
    }
}