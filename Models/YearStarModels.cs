using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models;

// 歲星神煞表：流年地支(1-12) × 宮位地支(1-12) → 吉星/凶星名稱
// YearId: 流年地支序號（子=1,丑=2,...,亥=12）
// FlowId: 宮位地支序號（子=1,丑=2,...,亥=12）
[Table("YearStarMap")]
public class YearStarMap
{
    [Key]
    public int Id { get; set; }

    public int YearId { get; set; }   // 流年地支序號 1-12

    public int FlowId { get; set; }   // 宮位地支序號 1-12

    [MaxLength(100)]
    public string GoodStar { get; set; } = "";  // 吉星（天德/將星/驛馬...）

    [MaxLength(200)]
    public string BadStar { get; set; } = "";   // 凶星（太歲/白虎/喪門...）
}

// 歲星年神論斷表：流年地支(1-12) × 宮位地支(1-12) → 年神名稱+論斷詩
// 十二年神：太歲/青龍/喪門/六合/官符/小耗/大耗/朱雀/白虎/貴神/吊客/病符
[Table("YearFlowStar")]
public class YearFlowStar
{
    [Key]
    public int Id { get; set; }

    public int YearId { get; set; }   // 流年地支序號 1-12

    public int FlowId { get; set; }   // 宮位地支序號 1-12

    [MaxLength(20)]
    public string StarName { get; set; } = "";  // 年神名稱（太歲/青龍...）

    [MaxLength(500)]
    public string Desc { get; set; } = "";      // 論斷詩

    [MaxLength(100)]
    public string StarType { get; set; } = "";  // 影響類型（貴人,事業,婚姻...）
}
