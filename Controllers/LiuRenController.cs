using Microsoft.AspNetCore.Mvc;
using Ecanapi.Data;
using Microsoft.EntityFrameworkCore;

namespace Ecanapi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class LiuRenController : ControllerBase
    {
        private readonly CalendarDbContext _calendarDb;

        private static readonly string[] _diZhi =
            ["子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥"];

        private static readonly string[] _tianGan =
            ["甲","乙","丙","丁","戊","己","庚","辛","壬","癸"];

        // 地支五行: 0=水 1=木 2=火 3=土 4=金
        private static readonly int[] _zhiElem =
            [0, 3, 1, 1, 3, 2, 2, 3, 4, 4, 3, 0];
        //   子 丑  寅  卯 辰  巳  午 未  申  酉 戌  亥

        // 月將名稱 (依地支順序 子~亥)
        private static readonly string[] _yueJiangName =
            ["神后","大吉","功曹","太衝","天罡","太乙","勝光","小吉","傳送","從魁","河魁","登明"];

        // 節氣 → 月將地支 index (月將在中氣後換，節沿用前中氣月將)
        private static readonly Dictionary<string, int> _jqYj = new()
        {
            {"冬至",1},{"大雪",1},   // 丑
            {"大寒",0},{"小寒",0},   // 子
            {"雨水",11},{"立春",11}, // 亥
            {"春分",10},{"驚蟄",10}, // 戌
            {"穀雨",9},{"清明",9},   // 酉
            {"小滿",8},{"立夏",8},   // 申
            {"夏至",7},{"芒種",7},   // 未
            {"大暑",6},{"小暑",6},   // 午
            {"處暑",5},{"立秋",5},   // 巳
            {"秋分",4},{"白露",4},   // 辰
            {"霜降",3},{"寒露",3},   // 卯
            {"小雪",2},{"立冬",2},   // 寅
        };

        // 十干寄宮 (甲→寅=2, 乙→辰=4, 丙→巳=5, 丁→未=7, 戊→巳=5,
        //           己→未=7, 庚→申=8, 辛→戌=10, 壬→亥=11, 癸→丑=1)
        private static readonly int[] _ganJiGong =
            [2, 4, 5, 7, 5, 7, 8, 10, 11, 1];

        // 陽貴人 (晝): 甲/戊/庚→丑=1, 乙/己→子=0, 丙/丁→亥=11, 壬/癸→卯=3, 辛→午=6
        private static readonly int[] _yangGui =
            [1, 0, 11, 11, 1, 0, 1, 6, 3, 3];

        // 陰貴人 (夜): 甲/戊/庚→未=7, 乙/己→申=8, 丙/丁→酉=9, 壬/癸→巳=5, 辛→午=6
        private static readonly int[] _yinGui =
            [7, 8, 9, 9, 7, 8, 7, 6, 5, 5];

        // 十二天將 (從貴人起)
        private static readonly string[] _tianJiang =
            ["貴人","螣蛇","朱雀","六合","勾陳","青龍","天空","白虎","太常","玄武","太陰","天后"];

        private static readonly Dictionary<string, string> _s2t = new()
        {
            {"小满","小滿"},{"处暑","處暑"},{"芒种","芒種"},
            {"谷雨","穀雨"},{"惊蛰","驚蟄"}
        };

        // 五行相克: 木(1)克土(3), 土(3)克水(0), 水(0)克火(2), 火(2)克金(4), 金(4)克木(1)
        private static bool Ke(int a, int b) => (a, b) is (1,3) or (3,0) or (0,2) or (2,4) or (4,1);

        public LiuRenController(CalendarDbContext calendarDb) => _calendarDb = calendarDb;

        [HttpGet("paipan")]
        public async Task<IActionResult> PaiPan(
            [FromQuery] string date,
            [FromQuery] int hour = 12)
        {
            if (!DateOnly.TryParse(date, out var d))
                return BadRequest("日期格式錯誤，請使用 yyyy-MM-dd");

            // ── 1. 曆法資料 ──────────────────────────────────────
            var cal = await _calendarDb.CalendarEntries.FirstOrDefaultAsync(
                c => c.Year == d.Year && c.SolarMonth == d.Month && c.SolarDay == d.Day);
            if (cal == null) return NotFound($"日期 {date} 不在曆法資料範圍");

            // ── 2. 節氣 → 月將 ───────────────────────────────────
            string stRaw = cal.SolarTerm ?? "";
            if (string.IsNullOrEmpty(stRaw)) return BadRequest("無節氣資料");
            int diPos = stRaw.IndexOf('第');
            string jieQiSimp = diPos > 0 ? stRaw[..diPos] : stRaw;
            string jieQi = _s2t.GetValueOrDefault(jieQiSimp, jieQiSimp);

            if (!_jqYj.TryGetValue(jieQi, out int yjIdx))
                return BadRequest($"未知節氣: {jieQi}");

            // ── 3. 占時地支 ──────────────────────────────────────
            int hourZhiIdx = hour == 23 ? 0 : (hour + 1) / 2;
            string hourZhi = _diZhi[hourZhiIdx];

            // ── 4. 日干支 ────────────────────────────────────────
            string gz = cal.DayGanzhi ?? "";
            if (gz.Length < 2) return BadRequest("日干支資料異常");
            string dayStem   = gz[..1];
            string dayBranch = gz[1..];
            int stemIdx     = Array.IndexOf(_tianGan, dayStem);
            int dayZhiIdx   = Array.IndexOf(_diZhi, dayBranch);

            // ── 5. 天盤配置: 月將加占時 ─────────────────────────
            // 天盤[地盤i] = (月將idx + i - 占時idx + 12) % 12
            int[] tp = new int[12];
            for (int i = 0; i < 12; i++)
                tp[i] = (yjIdx + i - hourZhiIdx + 12) % 12;

            // ── 6. 四課 ─────────────────────────────────────────
            // 課1: 下=日干(stem), 上=天盤[日干寄宮]
            int jiGong = _ganJiGong[stemIdx];
            int k1T = tp[jiGong]; int k1D_idx = jiGong; // 天/地 branch idx

            // 課2: 下=課1上, 上=天盤[課1上]
            int k2D = k1T; int k2T = tp[k1T];

            // 課3: 下=日支, 上=天盤[日支]
            int k3D = dayZhiIdx; int k3T = tp[dayZhiIdx];

            // 課4: 下=課3上, 上=天盤[課3上]
            int k4D = k3T; int k4T = tp[k3T];

            // ── 7. 三傳 (賊剋法) ─────────────────────────────────
            bool isYang = stemIdx % 2 == 0; // 甲丙戊庚壬=陽

            // 每課 (天branch_idx, 地branch_idx)；課1地=寄宮
            var lessons = new[] {
                (t: k1T, d: k1D_idx),
                (t: k2T, d: k2D),
                (t: k3T, d: k3D),
                (t: k4T, d: k4D),
            };

            int? chuFirst = null;
            string faName  = "賊剋";

            if (isYang)
            {
                // 陽日先取上克下 (天克地, 初傳=天課)
                foreach (var lk in lessons)
                    if (Ke(_zhiElem[lk.t], _zhiElem[lk.d])) { chuFirst = lk.t; break; }
                // 無則取下克上
                if (!chuFirst.HasValue)
                    foreach (var lk in lessons)
                        if (Ke(_zhiElem[lk.d], _zhiElem[lk.t])) { chuFirst = lk.t; faName = "賊剋(下克上)"; break; }
            }
            else
            {
                // 陰日先取下克上 (地克天, 初傳=天課)
                foreach (var lk in lessons)
                    if (Ke(_zhiElem[lk.d], _zhiElem[lk.t])) { chuFirst = lk.t; break; }
                if (!chuFirst.HasValue)
                    foreach (var lk in lessons)
                        if (Ke(_zhiElem[lk.t], _zhiElem[lk.d])) { chuFirst = lk.t; faName = "賊剋(上克下)"; break; }
            }

            // 比用法 fallback
            if (!chuFirst.HasValue)
            {
                faName    = "比用";
                chuFirst  = k1T; // 簡化: 取第一課天課
            }

            int sFirst  = chuFirst.Value;
            int sMiddle = tp[sFirst];
            int sLast   = tp[sMiddle];

            // ── 8. 十二天將 ──────────────────────────────────────
            // 晝 = 卯(3)~酉(9), 夜 = 其他
            bool isDay = hourZhiIdx >= 3 && hourZhiIdx <= 9;
            int guiBranch = isDay ? _yangGui[stemIdx] : _yinGui[stemIdx];

            // 找貴人支在天盤的地盤位置
            int guiPos = 0;
            for (int i = 0; i < 12; i++) if (tp[i] == guiBranch) { guiPos = i; break; }

            // 晝順(CW), 夜逆(CCW)
            var jPan = new string[12];
            for (int j = 0; j < 12; j++)
            {
                int pos = isDay ? (guiPos + j) % 12 : (guiPos - j + 12) % 12;
                jPan[pos] = _tianJiang[j];
            }

            // ── 9. 組裝 ─────────────────────────────────────────
            var positions = Enumerable.Range(0, 12).Select(i => (object)new
            {
                idx          = i,
                diPan        = _diZhi[i],
                tianPan      = _diZhi[tp[i]],
                tianJiang    = jPan[i],
                isHourBranch = i == hourZhiIdx,
                isDayBranch  = i == dayZhiIdx,
                isChuChuan   = tp[i] == sFirst,
                isZhongChuan = tp[i] == sMiddle,
                isMoChuan    = tp[i] == sLast,
            }).ToList();

            return Ok(new
            {
                date, hour,
                hourZhi,
                dayGanZhi   = $"{gz}日",
                dayStem, dayBranch, dayZhiIdx,
                jieQi,
                yueJiang     = _diZhi[yjIdx],
                yueJiangName = _yueJiangName[yjIdx],
                isDay        = isDay ? "晝" : "夜",
                guiRen       = _diZhi[guiBranch],
                siKe = new[]
                {
                    new { label="第一課", tian=_diZhi[k1T], di=dayStem,         diNote="日干" },
                    new { label="第二課", tian=_diZhi[k2T], di=_diZhi[k2D],     diNote="" },
                    new { label="第三課", tian=_diZhi[k3T], di=dayBranch,        diNote="日支" },
                    new { label="第四課", tian=_diZhi[k4T], di=_diZhi[k4D],     diNote="" },
                },
                sanChuan = new
                {
                    fa     = faName,
                    chu    = _diZhi[sFirst],
                    zhong  = _diZhi[sMiddle],
                    mo     = _diZhi[sLast],
                },
                positions,
            });
        }
    }
}
