using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models;

// 命宮星名表：12 地支 × 星名 + 命宮意涵（八字十二宮）
[Table("BaziMingGongStars")]
public class BaziMingGongStar
{
    [Key]
    public int Id { get; set; }

    // 命宮地支（子/丑/寅...亥）
    [MaxLength(4)]
    public string Branch { get; set; } = "";

    // 星名（天貴星/天厄星...）
    [MaxLength(20)]
    public string StarName { get; set; } = "";

    // 吉凶等級（大吉/吉/平/凶/大凶）
    [MaxLength(10)]
    public string LuckLevel { get; set; } = "";

    // 命宮論斷（志氣不凡，富裕清吉...）
    [MaxLength(200)]
    public string Description { get; set; } = "";
}

// 十二宮神煞表：流年/流月神煞 × 說明（流年命書用）
[Table("BaziShenSha12")]
public class BaziShenSha12
{
    [Key]
    public int Id { get; set; }

    // 神煞在十二宮中的順序（1=太歲, 2=太陽...12=百越）
    public int Sequence { get; set; }

    // 神煞名稱（太歲/太陽/驛馬...）
    [MaxLength(20)]
    public string ShenShaName { get; set; } = "";

    // 吉凶等級（大吉/吉/中性/小凶/凶/大凶）
    [MaxLength(10)]
    public string LuckLevel { get; set; } = "";

    // 簡短說明（命書呈現用）
    [MaxLength(100)]
    public string ShortDesc { get; set; } = "";

    // 完整說明（備用/詳解）
    [MaxLength(500)]
    public string FullDesc { get; set; } = "";
}
