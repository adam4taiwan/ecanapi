using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models
{
    [Table("BaziDirectRules")]
    public class BaziDirectRule
    {
        [Key]
        public int Id { get; set; }

        // Chapter 1~12
        public int Chapter { get; set; }

        // Section within chapter (0 = general)
        public int Section { get; set; }

        // Rule category: HourBranch, HourStem, StemRepeat, BranchRepeat,
        // ParentInfo, SiblingInfo, ChildInfo, Marriage, Career,
        // TenGod, Death, Injury, BodyTrait, JianghuSecret, DaYun
        [MaxLength(50)]
        public string RuleType { get; set; } = "";

        // Trigger condition (e.g. "子午卯酉", "甲乙時干", "三子")
        [MaxLength(200)]
        public string Condition { get; set; } = "";

        // The reading content
        public string Content { get; set; } = "";

        // Supplementary notes or source verse
        public string? Notes { get; set; }

        public int SortOrder { get; set; }

        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }
}
