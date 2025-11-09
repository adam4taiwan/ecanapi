// 檔案：Models/Analysis/NaYin.cs
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models.Analysis
{
    [Table("納音")]
    public class NaYin
    {
        [Key]
        [Column("unique_id")]
        public int UniqueId { get; set; }

        [Column("gan_zhi")]
        public string? GanZhi { get; set; }

        [Column("na_yin")]
        public string? NaYinValue { get; set; }
    }
}