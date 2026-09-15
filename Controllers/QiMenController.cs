using Microsoft.AspNetCore.Mvc;
using Ecanapi.Data;
using Ecanapi.Models;
using Microsoft.EntityFrameworkCore;

namespace Ecanapi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class QiMenController : ControllerBase
    {
        private readonly ApplicationDbContext _context;
        private readonly CalendarDbContext _calendarDb;

        // ============================================================
        //  Static reference tables
        // ============================================================

        private static readonly string[] _tianGan =
            ["甲","乙","丙","丁","戊","己","庚","辛","壬","癸"];

        private static readonly string[] _diZhi =
            ["子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥"];

        // 六旬首六儀: 甲子旬→戊, 甲戌→己, 甲申→庚, 甲午→辛, 甲辰→壬, 甲寅→癸
        private static readonly string[] _liuYi =
            ["戊","己","庚","辛","壬","癸"];

        // 九星 (SortOrder 1-9, 對應宮位 1-9)
        private static readonly string[] _jiuXing =
            ["天蓬","天芮","天沖","天輔","天禽","天心","天柱","天任","天英"];

        // 八門 (SortOrder 1-8), 本宮: 1坎,2坤,3震,4巽,6乾,7兌,8艮,9離
        private static readonly string[] _baMen =
            ["休門","死門","傷門","杜門","開門","驚門","生門","景門"];
        private static readonly int[] _baMenOrigPalace = [1,2,3,4,6,7,8,9];

        // 八神 (順序: 值符→騰蛇→太陰→六合→白虎→玄武→九地→九天)
        private static readonly string[] _baShen =
            ["值符","騰蛇","太陰","六合","白虎","玄武","九地","九天"];

        private static readonly Dictionary<int,string> _gongName = new()
            { {1,"坎"},{2,"坤"},{3,"震"},{4,"巽"},{5,"中"},{6,"乾"},{7,"兌"},{8,"艮"},{9,"離"} };

        // 洛書九宮在 3×3 格的位置 (row, col)
        private static readonly Dictionary<int,(int row,int col)> _gongPos = new()
        {
            {4,(0,0)},{9,(0,1)},{2,(0,2)},
            {3,(1,0)},{5,(1,1)},{7,(1,2)},
            {8,(2,0)},{1,(2,1)},{6,(2,2)}
        };

        // 八宮順時針外圈 (驗證: 從 7兌起 7→6→1→8→3→4→9→2)
        private static readonly int[] _cwOuter = [4,9,2,7,6,1,8,3];

        // 簡體→繁體節氣對照 (calendar 表用簡體)
        private static readonly Dictionary<string,string> _s2t = new()
        {
            {"小满","小滿"},{"处暑","處暑"},{"芒种","芒種"},
            {"谷雨","穀雨"},{"惊蛰","驚蟄"}
        };

        // 陰遁節氣集合 (繁體)
        private static readonly HashSet<string> _yinJieQi = new()
        {
            "夏至","小暑","大暑","立秋","處暑","白露",
            "秋分","寒露","霜降","立冬","小雪","大雪"
        };

        // 五鼠遁: 甲己→甲子(0), 乙庚→丙子(2), 丙辛→戊子(4), 丁壬→庚子(6), 戊癸→壬子(8)
        private static readonly int[] _ziStemStart = [0,2,4,6,8];

        // ============================================================
        //  Constructor
        // ============================================================

        public QiMenController(ApplicationDbContext context, CalendarDbContext calendarDb)
        {
            _context = context;
            _calendarDb = calendarDb;
        }

        // ============================================================
        //  GET /api/QiMen/paipan?date=2024-09-15&hour=14
        // ============================================================

        /// <summary>奇門遁甲排盤（妙派標準）</summary>
        [HttpGet("paipan")]
        public async Task<IActionResult> PaiPan(
            [FromQuery] string date,
            [FromQuery] int hour = 12)
        {
            if (!DateOnly.TryParse(date, out var d))
                return BadRequest("日期格式錯誤，請使用 yyyy-MM-dd");

            // ── 1. 取得曆法資料 ──────────────────────────────────────
            var cal = await _calendarDb.CalendarEntries.FirstOrDefaultAsync(
                c => c.Year == d.Year && c.SolarMonth == d.Month && c.SolarDay == d.Day);
            if (cal == null)
                return NotFound($"日期 {date} 不在曆法資料範圍內");

            // ── 2. 解析節氣 ──────────────────────────────────────────
            string stRaw = cal.SolarTerm ?? "";
            if (string.IsNullOrEmpty(stRaw))
                return BadRequest("無法取得節氣資料，請確認曆法資料是否完整");

            int diPos   = stRaw.IndexOf('第');
            int tianPos = stRaw.IndexOf('天');
            if (diPos < 0 || tianPos < 0)
                return BadRequest($"節氣格式異常: {stRaw}");

            string jieQiSimp = stRaw[..diPos];
            int    jieQiDay  = int.Parse(stRaw[(diPos + 1)..tianPos]);
            string jieQi     = _s2t.GetValueOrDefault(jieQiSimp, jieQiSimp);

            // ── 3. 元 / 陰陽遁 ───────────────────────────────────────
            string yuan    = jieQiDay <= 5 ? "上元" : jieQiDay <= 10 ? "中元" : "下元";
            bool   isYin   = _yinJieQi.Contains(jieQi);
            string yinYang = isYin ? "陰遁" : "陽遁";

            // ── 4. 查局數 ────────────────────────────────────────────
            var juCfg = await _context.QiMenJuConfigs.FirstOrDefaultAsync(
                q => q.YinYang == yinYang && q.JieQi == jieQi && q.Yuan == yuan);
            if (juCfg == null)
                return NotFound($"局配置不存在: {yinYang} {jieQi} {yuan}");
            int juShu = juCfg.JuShu;

            // ── 5. 取地盤天干 (9 宮) ─────────────────────────────────
            var diPanList = await _context.QiMenDiPanMaps
                .Where(m => m.YinYang == yinYang && m.JuShu == juShu)
                .ToListAsync();
            if (diPanList.Count < 9)
                return StatusCode(500, "地盤資料不完整");
            var diPan = diPanList.ToDictionary(m => m.GongWei, m => m.TianGan);

            // ── 6. 日干支 ────────────────────────────────────────────
            string gz = cal.DayGanzhi ?? "";
            if (gz.Length < 2) return BadRequest("日干支資料異常");
            string dayStem   = gz[..1];
            string dayBranch = gz[1..];

            // ── 7. 時支 / 時干 (五鼠遁) ─────────────────────────────
            // 子時: 23:00-00:59, 丑時: 01-02, 寅時: 03-04 ...
            int hourBranchIdx = hour == 23 ? 0 : (hour + 1) / 2;
            string hourBranch = _diZhi[hourBranchIdx];

            int dStemIdx = Array.IndexOf(_tianGan, dayStem);
            int ziStem   = _ziStemStart[dStemIdx % 5];
            string hourStem = _tianGan[(ziStem + hourBranchIdx) % 10];

            // ── 8. 旬首六儀 ──────────────────────────────────────────
            // 利用 CRT: p60 = sIdx + 10 * (diffBranch/2 * 5 % 6)
            int sIdx      = Array.IndexOf(_tianGan, dayStem);
            int bIdx      = Array.IndexOf(_diZhi, dayBranch);
            int diffBranch = (bIdx - sIdx + 12) % 12; // 應為偶數
            int p60        = sIdx + 10 * (diffBranch / 2 * 5 % 6);
            int xunIdx     = (p60 / 10) % 6;
            string liuYi   = _liuYi[xunIdx];

            // ── 9. P (旬首宮) / Q (時干宮) ───────────────────────────
            // 中宮(5)一律寄坤二(2)
            int P     = diPan.First(kv => kv.Value == liuYi).Key;
            int Q_raw = diPan.First(kv => kv.Value == hourStem).Key;
            int Q     = Q_raw == 5 ? 2 : Q_raw;
            int P_eff = P == 5 ? 2 : P;  // 天禽在中宮時寄坤二

            string valueFuXing = _jiuXing[P - 1]; // 旬首宮本星 = 值符天星

            // ── 10. 天盤 (九星) ──────────────────────────────────────
            // 陽遁: 星 CW 位移; 陰遁: 星 CCW 位移 (= 負 CW 位移)
            // 公式: newCwIdx = (origCwIdx ± shift) % 8
            int pCwIdx = Array.IndexOf(_cwOuter, P_eff);
            int qCwIdx = Array.IndexOf(_cwOuter, Q);
            var tianXingPan = new Dictionary<int,string>();
            if (isYin)
            {
                int ccwShift = (pCwIdx - qCwIdx + 8) % 8; // CCW 步數
                foreach (int outerP in _cwOuter)
                {
                    int oIdx   = Array.IndexOf(_cwOuter, outerP);
                    int newIdx = (oIdx - ccwShift + 8) % 8;
                    tianXingPan[_cwOuter[newIdx]] = _jiuXing[outerP - 1];
                }
            }
            else
            {
                int cwShift = (qCwIdx - pCwIdx + 8) % 8; // CW 步數
                foreach (int outerP in _cwOuter)
                {
                    int oIdx   = Array.IndexOf(_cwOuter, outerP);
                    int newIdx = (oIdx + cwShift) % 8;
                    tianXingPan[_cwOuter[newIdx]] = _jiuXing[outerP - 1];
                }
            }
            tianXingPan[5] = "天禽"; // 天禽永遠在中宮

            // ── 11. 值使宮 R (排山掌) ────────────────────────────────
            // 從旬首宮 P_eff 起甲(1步), 陽順陰逆數至時干位置
            // 時干位置 = Array.IndexOf(_tianGan, hourStem) + 1  (甲=1,...,癸=10)
            int hourStemPos = Array.IndexOf(_tianGan, hourStem) + 1; // 1-10
            int steps       = hourStemPos - 1; // 移動步數 (0=停在起點)
            int R_raw;
            if (isYin)
                R_raw = ((P_eff - 1 - steps + 9 * 1000) % 9) + 1;
            else
                R_raw = ((P_eff - 1 + steps) % 9) + 1;
            int R = R_raw == 5 ? 2 : R_raw; // 中宮寄坤二

            // ── 12. 八門盤 (永遠順時針旋轉) ─────────────────────────
            // 值使門 = 旬首宮 P_eff 的本門 (SortOrder)
            int valueShuiMenSO  = BaMenSOForPalace(P_eff);
            string valueShuiMen = _baMen[valueShuiMenSO - 1];
            int origMenPalace   = _baMenOrigPalace[valueShuiMenSO - 1];
            int origMenCwIdx    = Array.IndexOf(_cwOuter, origMenPalace);
            int rCwIdx          = Array.IndexOf(_cwOuter, R);
            int baMenCwOffset   = (rCwIdx - origMenCwIdx + 8) % 8;
            var baMenPan = new Dictionary<int,string>();
            for (int i = 0; i < 8; i++)
            {
                int oIdx   = Array.IndexOf(_cwOuter, _baMenOrigPalace[i]);
                int newIdx = (oIdx + baMenCwOffset) % 8;
                baMenPan[_cwOuter[newIdx]] = _baMen[i];
            }

            // ── 13. 八神盤 ──────────────────────────────────────────
            // 值符置於 Q 宮; 陽順(CW) 陰逆(CCW) 排其餘七神
            var baShenPan = new Dictionary<int,string>();
            for (int i = 0; i < 8; i++)
            {
                int palace = isYin
                    ? _cwOuter[(qCwIdx - i + 80) % 8]
                    : _cwOuter[(qCwIdx + i) % 8];
                baShenPan[palace] = _baShen[i];
            }

            // ── 14. 組裝 9 宮結果 ────────────────────────────────────
            var palaces = new List<object>();
            for (int g = 1; g <= 9; g++)
            {
                var pos = _gongPos[g];
                string xing  = tianXingPan.GetValueOrDefault(g, "");
                string men   = baMenPan.ContainsKey(g) ? baMenPan[g]   : (g == 5 ? "無門" : "");
                string shen  = baShenPan.ContainsKey(g) ? baShenPan[g]  : "";
                palaces.Add(new
                {
                    gong        = g,
                    gongName    = _gongName[g],
                    row         = pos.row,
                    col         = pos.col,
                    diPanStem   = diPan.GetValueOrDefault(g, ""),
                    tianXing    = xing,
                    baMen       = men,
                    baShen      = shen,
                    isValueFu   = (g == Q),    // 值符宮 (時干宮)
                    isValueShui = (g == R)     // 值使宮
                });
            }

            return Ok(new
            {
                date          = date,
                hour          = hour,
                hourBranch    = hourBranch,
                hourStem      = hourStem,
                hourGanZhi    = $"{hourStem}{hourBranch}時",
                dayGanZhi     = $"{gz}日",
                jieQi         = jieQi,
                jieQiDay      = jieQiDay,
                yuan          = yuan,
                yinYang       = yinYang,
                juShu         = juShu,
                xunShouLiuYi  = liuYi,
                xunShouGong   = P,
                valueFuXing   = valueFuXing,
                valueFuGong   = Q,
                valueShuiMen  = valueShuiMen,
                valueShuiGong = R,
                palaces       = palaces
            });
        }

        // ============================================================
        //  Helper
        // ============================================================

        /// <summary>取宮位對應的八門 SortOrder (1-8); 中宮(5)→死門(2)</summary>
        private static int BaMenSOForPalace(int palace) => palace switch
        {
            1 => 1, 2 => 2, 3 => 3, 4 => 4,
            5 => 2, // 禽星無門, 以死門代值使
            6 => 5, 7 => 6, 8 => 7, 9 => 8,
            _ => 1
        };
    }
}
