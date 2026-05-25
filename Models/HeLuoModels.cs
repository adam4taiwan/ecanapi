using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models
{
    [Table("ig", Schema = "public")]
    public class IgHexagram
    {
        [Key]
        [Column("id")]
        public int Id { get; set; }

        [Column("code")]
        public string Code { get; set; } = "";

        [Column("name")]
        public string? Name { get; set; }

        [Column("description")]
        public string? Description { get; set; }

        [Column("desc_one")]
        public string? DescOne { get; set; }

        [Column("desc_two")]
        public string? DescTwo { get; set; }

        [Column("desc_three")]
        public string? DescThree { get; set; }

        [Column("desc_four")]
        public string? DescFour { get; set; }

        [Column("desc_five")]
        public string? DescFive { get; set; }

        [Column("desc_six")]
        public string? DescSix { get; set; }
    }

    [Table("ig64_six", Schema = "public")]
    public class Ig64Six
    {
        [Key]
        [Column("id")]
        public int Id { get; set; }

        [Column("ig64")]
        public string? Ig64 { get; set; }

        [Column("one_yao")]
        public string? OneYao { get; set; }

        [Column("two_yao")]
        public string? TwoYao { get; set; }

        [Column("three_yao")]
        public string? ThreeYao { get; set; }

        [Column("four_yao")]
        public string? FourYao { get; set; }

        [Column("five_yao")]
        public string? FiveYao { get; set; }

        [Column("six_yao")]
        public string? SixYao { get; set; }

        [Column("RowID")]
        public int RowId { get; set; }

        [Column("gongming")]
        public string? Gongming { get; set; }

        [Column("wuxing")]
        public string? Wuxing { get; set; }
    }
}
