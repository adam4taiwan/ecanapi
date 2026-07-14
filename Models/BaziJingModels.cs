using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models
{
    // Ch.3 - 命局氣勢格局高低（吉凶組合）
    [Table("BaziJingConfig")]
    public class BaziJingConfig
    {
        [Key]
        [Column("Id")]
        public int Id { get; set; }

        // 吉 or 凶
        [Column("ConfigType")]
        public string ConfigType { get; set; } = "";

        // e.g. 身財兩停, 七殺攻身
        [Column("ConfigName")]
        public string ConfigName { get; set; } = "";

        // keyword for code logic, e.g. "SHEN_CAI_LIANG_TING"
        [Column("Condition")]
        public string? Condition { get; set; }

        [Column("Content")]
        public string? Content { get; set; }

        // Practitioner-authored advice (A part of C+A display)
        [Column("AdviceText")]
        public string? AdviceText { get; set; }

        [Column("SortOrder")]
        public int SortOrder { get; set; }
    }

    // Ch.4 - 財官論命
    [Table("BaziJingCaiGuan")]
    public class BaziJingCaiGuan
    {
        [Key]
        [Column("Id")]
        public int Id { get; set; }

        // 財 / 官 / 互動
        [Column("Category")]
        public string Category { get; set; } = "";

        // e.g. 比劫取財, 食傷生財, 官印相生
        [Column("ConfigType")]
        public string ConfigType { get; set; } = "";

        [Column("Condition")]
        public string? Condition { get; set; }

        [Column("Content")]
        public string? Content { get; set; }

        // Practitioner-authored advice (A part of C+A display)
        [Column("AdviceText")]
        public string? AdviceText { get; set; }

        [Column("SortOrder")]
        public int SortOrder { get; set; }
    }

    // Ch.5 - 干支象法（10天干 + 12地支）
    [Table("BaziJingXiang")]
    public class BaziJingXiang
    {
        [Key]
        [Column("Id")]
        public int Id { get; set; }

        // 天干 or 地支
        [Column("XiangType")]
        public string XiangType { get; set; } = "";

        // 甲乙丙... / 子丑寅...
        [Column("Key")]
        public string Key { get; set; } = "";

        [Column("BasicImage")]
        public string? BasicImage { get; set; }

        [Column("BodyImage")]
        public string? BodyImage { get; set; }

        [Column("PersonImage")]
        public string? PersonImage { get; set; }

        [Column("CareerImage")]
        public string? CareerImage { get; set; }

        [Column("RelationImage")]
        public string? RelationImage { get; set; }

        [Column("Notes")]
        public string? Notes { get; set; }
    }

    // Ch.6 - 神煞（查法 + 論斷）
    [Table("BaziJingShenSha")]
    public class BaziJingShenSha
    {
        [Key]
        [Column("Id")]
        public int Id { get; set; }

        [Column("Name")]
        public string Name { get; set; } = "";

        // 年支 / 日干 / 日支
        [Column("LookupBase")]
        public string LookupBase { get; set; } = "";

        // JSON: {"子":"寅","丑":"亥",...}
        [Column("LookupMap")]
        public string? LookupMap { get; set; }

        [Column("AuspiciousText")]
        public string? AuspiciousText { get; set; }

        [Column("InauspiciousText")]
        public string? InauspiciousText { get; set; }

        [Column("SpecialRule")]
        public string? SpecialRule { get; set; }

        [Column("SortOrder")]
        public int SortOrder { get; set; }
    }

    // Ch.7 - 盲派口訣（年/月/日/時柱口訣 + 十排歌 + 十神口訣）
    [Table("BaziJingKouJue")]
    public class BaziJingKouJue
    {
        [Key]
        [Column("Id")]
        public int Id { get; set; }

        // 年柱 / 月柱 / 日柱 / 時柱 / 十排歌 / 十神口訣 / 兩神組合
        [Column("Category")]
        public string Category { get; set; } = "";

        // 條件鍵，e.g. "年干=甲", "月支=子", "十神=殺_重=1"
        [Column("Condition")]
        public string? Condition { get; set; }

        [Column("Content")]
        public string? Content { get; set; }

        [Column("SortOrder")]
        public int SortOrder { get; set; }
    }

    // Ch.8 - 六親論斷
    [Table("BaziJingLiuQin")]
    public class BaziJingLiuQin
    {
        [Key]
        [Column("Id")]
        public int Id { get; set; }

        // 父 / 母 / 兄弟 / 配偶 / 子女
        [Column("LiuQinType")]
        public string LiuQinType { get; set; } = "";

        // 個數 / 時機 / 品質 / 克損
        [Column("Category")]
        public string Category { get; set; } = "";

        [Column("Condition")]
        public string? Condition { get; set; }

        [Column("Content")]
        public string? Content { get; set; }

        [Column("SortOrder")]
        public int SortOrder { get; set; }
    }

    // Ch.9/10 - 歲運論斷
    [Table("BaziJingYunShi")]
    public class BaziJingYunShi
    {
        [Key]
        [Column("Id")]
        public int Id { get; set; }

        // 大運 / 流年 / 共通
        [Column("Category")]
        public string Category { get; set; } = "";

        // 短標題 e.g. 喜用運, 忌神運, 財運
        [Column("Title")]
        public string? Title { get; set; }

        // 詳細條件描述 e.g. 逢喜用神大運或流年
        [Column("Condition")]
        public string? Condition { get; set; }

        [Column("Content")]
        public string? Content { get; set; }

        [Column("SortOrder")]
        public int SortOrder { get; set; }
    }
}
