// 檔案：Models/Analysis/IChing64Hexagrams.cs
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models.Analysis
{
    [Table("易經六十四卦")]
    public class IChing64Hexagrams
    {
        [Key]
        [Column("gua_id")]
        public int GuaId { get; set; }

        [Column("gua_value")]
        public int? GuaValue { get; set; }

        [Column("gua_name")]
        public string? GuaName { get; set; }

        [Column("gua_instruction")]
        public string? GuaInstruction { get; set; }
    }
}