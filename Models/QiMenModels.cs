using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models
{
    // 九星基本資料（天蓬/天芮/天沖/天輔/天禽/天心/天柱/天任/天英）
    public class QiMenJiuXing
    {
        [Key]
        public int Id { get; set; }
        public string Name { get; set; } = "";           // 天蓬/天芮/...
        public int DefaultGong { get; set; }             // 固定本宮（洛書數 1-9）
        public string WuXing { get; set; } = "";         // 木/火/土/金/水
        public string JiXiong { get; set; } = "";        // 大吉/吉/中/凶/大凶
        public string Desc { get; set; } = "";           // 基本含義
        public int SortOrder { get; set; }
    }

    // 八門基本資料（休/生/傷/杜/景/死/驚/開）
    public class QiMenBaMen
    {
        [Key]
        public int Id { get; set; }
        public string Name { get; set; } = "";           // 休/生/傷/杜/景/死/驚/開
        public int DefaultGong { get; set; }             // 固定本宮（洛書數 1-9）
        public string WuXing { get; set; } = "";
        public string JiXiong { get; set; } = "";        // 大吉/吉/中/凶/大凶
        public string Desc { get; set; } = "";
        public int SortOrder { get; set; }
    }

    // 八神基本資料（值符/騰蛇/太陰/六合/白虎/玄武/九地/九天）
    public class QiMenBaShen
    {
        [Key]
        public int Id { get; set; }
        public string Name { get; set; } = "";           // 值符/騰蛇/...
        public string JiXiong { get; set; } = "";        // 大吉/吉/中/凶/大凶
        public string Desc { get; set; } = "";
        public int SortOrder { get; set; }               // 固定順序（值符=1，後續依序）
    }

    // 六儀三奇天干屬性（戊己庚辛壬癸丁丙乙，地盤排列順序）
    public class QiMenTianGan
    {
        [Key]
        public int Id { get; set; }
        public string TianGan { get; set; } = "";        // 戊/己/庚/辛/壬/癸/丁/丙/乙
        public string Type { get; set; } = "";           // 六儀/三奇
        public string JiXiong { get; set; } = "";        // 大吉/吉/中/凶/大凶
        public string Desc { get; set; } = "";
        public int SortOrder { get; set; }               // 地盤排列順序（戊=1...乙=9）
    }

    // 地盤天干對照（陽/陰遁 × 局1-9 × 宮1-9 → 天干）
    // 陽遁順飛：戊起局數宮，按洛書 1→2→...→9→1 順序布
    // 陰遁逆飛：戊起局數宮，按洛書 9→8→...→1→9 逆序布
    public class QiMenDiPanMap
    {
        [Key]
        public int Id { get; set; }
        public string YinYang { get; set; } = "";        // 陽遁/陰遁
        public int JuShu { get; set; }                   // 局數 1-9
        public int GongWei { get; set; }                 // 宮位 1-9（洛書數）
        public string TianGan { get; set; } = "";        // 戊/己/庚/辛/壬/癸/丁/丙/乙
    }

    // 節氣元局對照（節氣 × 上/中/下元 → 局數）
    // 陽遁：冬至至芒種（12節氣）；陰遁：夏至至大雪（12節氣）
    // 注意：此表依標準法填入，玉洞子驗證後可按需調整
    public class QiMenJuConfig
    {
        [Key]
        public int Id { get; set; }
        public string YinYang { get; set; } = "";        // 陽遁/陰遁
        public string JieQi { get; set; } = "";          // 節氣名（冬至/小寒/...）
        public int JieQiOrder { get; set; }              // 節氣排序（陽1-12/陰13-24）
        public string Yuan { get; set; } = "";           // 上元/中元/下元
        public int JuShu { get; set; }                   // 局數 1-9
    }
}
