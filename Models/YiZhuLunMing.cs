using System.ComponentModel.DataAnnotations;

namespace Ecanapi.Models
{
    /// <summary>
    /// 一柱論命·六十甲子日柱定數
    /// </summary>
    public class YiZhuLunMing
    {
        [Key]
        public int Id { get; set; }

        /// <summary>日柱干支，如「甲子」</summary>
        [MaxLength(2)]
        public string DayPillar { get; set; } = "";

        /// <summary>性格特質（多段描述）</summary>
        public string? Personality { get; set; }

        /// <summary>四句詩/格言</summary>
        public string? Poem { get; set; }

        /// <summary>月令論斷，JSON 格式 {"子":"...","丑":"...",...}</summary>
        public string? MonthlyAnalysis { get; set; }

        /// <summary>空亡論斷（備用）</summary>
        public string? VoidAnalysis { get; set; }
    }
}
