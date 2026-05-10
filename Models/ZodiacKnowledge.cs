using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models
{
    [Table("生肖命理庫")]
    public class ZodiacKnowledge
    {
        [Key]
        [Column("id")]
        public int Id { get; set; }

        // 出生地支：子丑寅卯辰巳午未申酉戌亥
        [Column("birth_branch")]
        public string BirthBranch { get; set; } = "";

        // 出生生肖：鼠牛虎兔龍蛇馬羊猴雞狗豬
        [Column("birth_zodiac")]
        public string BirthZodiac { get; set; } = "";

        // 大分類：本命特性 / 精批榮祿 / 流年運程 / 流月運程 / 改名增運
        [Column("category")]
        public string Category { get; set; } = "";

        // 子分類：健康/財運/事業/職業/愛情/婚姻/守護神/名人堂/生年/生月/生時辰/取名 等
        [Column("subcategory")]
        public string? Subcategory { get; set; }

        // 流年用：逢X年的地支（子丑寅...）；流月用：月地支（寅=正月...）；其他 NULL
        [Column("target_branch")]
        public string? TargetBranch { get; set; }

        // 人類可讀標籤：「逢子年」「正月」「甲子年」等
        [Column("target_label")]
        public string? TargetLabel { get; set; }

        [Column("content")]
        public string Content { get; set; } = "";

        // 吉凶五類：大吉 / 吉 / 平 / 凶 / 大凶
        [Column("fortune_level")]
        public string? FortuneLevel { get; set; }

        [Column("uid")]
        public int Uid { get; set; }
    }
}
