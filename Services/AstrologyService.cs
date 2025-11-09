using Ecan;
using Ecanapi.Data;
using Ecanapi.Models;
using Ecanapi.Services.AstrologyEngine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ecanapi.Services
{
    public class AstrologyService : IAstrologyService
    {
        private readonly ICalendarService _calendarService;

        public AstrologyService(ICalendarService calendarService)
        {
            _calendarService = calendarService;
        }

        #region --- Bazi & Brightness Data ---
        private static readonly Dictionary<string, string> NaYinMap = new Dictionary<string, string> { { "甲子", "海中金" }, { "乙丑", "海中金" }, { "丙寅", "爐中火" }, { "丁卯", "爐中火" }, { "戊辰", "大林木" }, { "己巳", "大林木" }, { "庚午", "路旁土" }, { "辛未", "路旁土" }, { "壬申", "劍鋒金" }, { "癸酉", "劍鋒金" }, { "甲戌", "山頭火" }, { "乙亥", "山頭火" }, { "丙子", "澗下水" }, { "丁丑", "澗下水" }, { "戊寅", "城頭土" }, { "己卯", "城頭土" }, { "庚辰", "白蠟金" }, { "辛巳", "白蠟金" }, { "壬午", "楊柳木" }, { "癸未", "楊柳木" }, { "甲申", "泉中水" }, { "乙酉", "泉中水" }, { "丙戌", "屋上土" }, { "丁亥", "屋上土" }, { "戊子", "霹靂火" }, { "己丑", "霹靂火" }, { "庚寅", "松柏木" }, { "辛卯", "松柏木" }, { "壬辰", "長流水" }, { "癸巳", "長流水" }, { "甲午", "沙中金" }, { "乙未", "沙中金" }, { "丙申", "山下火" }, { "丁酉", "山下火" }, { "戊戌", "平地木" }, { "己亥", "平地木" }, { "庚子", "壁上土" }, { "辛丑", "壁上土" }, { "壬寅", "金箔金" }, { "癸卯", "金箔金" }, { "甲辰", "覆燈火" }, { "乙巳", "覆燈火" }, { "丙午", "天河水" }, { "丁未", "天河水" }, { "戊申", "大驛土" }, { "己酉", "大驛土" }, { "庚戌", "釵釧金" }, { "辛亥", "釵釧金" }, { "壬子", "桑柘木" }, { "癸丑", "桑柘木" }, { "甲寅", "大溪水" }, { "乙卯", "大溪水" }, { "丙辰", "沙中土" }, { "丁巳", "沙中土" }, { "戊午", "天上火" }, { "己未", "天上火" }, { "庚申", "石榴木" }, { "辛酉", "石榴木" }, { "壬戌", "大海水" }, { "癸亥", "大海水" } };
        private static readonly Dictionary<int, string[]> HiddenStemsMap = new Dictionary<int, string[]> { { 1, new[] { "癸" } }, { 2, new[] { "己", "癸", "辛" } }, { 3, new[] { "戊", "丙", "甲" } }, { 4, new[] { "乙" } }, { 5, new[] { "戊", "乙", "癸" } }, { 6, new[] { "庚", "丙", "戊" } }, { 7, new[] { "丁", "己" } }, { 8, new[] { "己", "丁", "乙" } }, { 9, new[] { "壬", "庚", "戊" } }, { 10, new[] { "辛" } }, { 11, new[] { "戊", "辛", "丁" } }, { 12, new[] { "壬", "甲" } } };
        private string GetLiuShen(string dayMasterGan, string otherGan) { if (dayMasterGan == otherGan) return "比"; int dayMasterIndex = "甲乙丙丁戊己庚辛壬癸".IndexOf(dayMasterGan); int otherIndex = "甲乙丙丁戊己庚辛壬癸".IndexOf(otherGan); int dayMasterWuXing = dayMasterIndex / 2; int otherWuXing = otherIndex / 2; bool isDayMasterYang = dayMasterIndex % 2 == 0; bool isOtherYang = otherIndex % 2 == 0; if (otherWuXing == (dayMasterWuXing + 1) % 5) return isDayMasterYang == isOtherYang ? "食" : "傷"; if (otherWuXing == (dayMasterWuXing + 2) % 5) return isDayMasterYang == isOtherYang ? "才" : "財"; if (otherWuXing == (dayMasterWuXing + 3) % 5) return isDayMasterYang == isOtherYang ? "殺" : "官"; if (otherWuXing == (dayMasterWuXing + 4) % 5) return isDayMasterYang == isOtherYang ? "梟" : "印"; return "劫"; }
        private static readonly string[] S_SKY = { "", "甲", "乙", "丙", "丁", "戊", "己", "庚", "辛", "壬", "癸" };
        private static readonly string[] S_FLOOR = { "", "子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥" };
        private static readonly string[] NUMERIC_LIGHT_TABLE = { "", "33323133113321", "22313033003300", "23213032003300", "22313032012300", "22313033002201", "32332233122222", "21213032003300", "32332333113313", "21322233103301", "32322322223222", "21232133012221", "21213133003310" };
        private static readonly string[] SYMBOL_LIGHT_TABLE = { "", "+*!+*!+*++!*++", "!!++!+*++!*++!", "+!+++++!!++++-", "!+*++-+!!+!*+-", "++*++++!+!++++", "+!*!+-+!-!+++!", "***+!!+!++!*++", "!!++!+*++!+++!", "++!+!+++!!+!+-", "!+!++-+*!+!!+-", "++!++++*+!++++", "+!!!+-+*-!+++!" };
        private static readonly string N12S = "命父兄夫子財疾遷奴官田福";
        private static readonly string M6S = "紫機陽武同廉";
        private static readonly string N8S = "府陰貪巨相梁殺破";
        #endregion

        public async Task<AstrologyChartResult> CalculateChartAsync(AstrologyRequest request)
        {
            var context = new AstrologyCalculationContext(request);
            await DetermineCorrectBaziPillars(context);

            DetermineBaziLuckCycles(context);

            RunZiWeiChartCalculations(context);
            DetermineMajorStars(context);
            DetermineAllAuxiliaryAndMinorStars(context);
            DetermineDecadeLuckCycles(context);
            DetermineLifeCycleStars(context);
            DetermineStarBrightness(context);
            UpdateFinalResult(context);
            return context.Result;
        }

        private async Task DetermineCorrectBaziPillars(AstrologyCalculationContext context)
        {
            var birthDate = context.Request;
            var calendar = context.Calendar;
            string yearGan, yearZhi, monthGan, monthZhi, dayGan, dayZhi, hourGan, hourZhi;

            var calendarData = await _calendarService.GetCalendarDataAsync(birthDate.Year, birthDate.Month, birthDate.Day);

            if (calendarData != null && !string.IsNullOrEmpty(calendarData.YearGanzhi))
            {
                yearGan = calendarData.YearGanzhi.Substring(0, 1);
                yearZhi = calendarData.YearGanzhi.Substring(1, 1);
                monthGan = calendarData.MonthGanzhi.Substring(0, 1);
                monthZhi = calendarData.MonthGanzhi.Substring(1, 1);
                dayGan = calendarData.DayGanzhi.Substring(0, 1);
                dayZhi = calendarData.DayGanzhi.Substring(1, 1);
            }
            else
            {
                yearGan = calendar.GanZhiYYString.Substring(0, 1);
                yearZhi = calendar.GanZhiYYString.Substring(1, 1);
                monthGan = calendar.GanZhiMMString.Substring(0, 1);
                monthZhi = calendar.GanZhiMMString.Substring(1, 1);
                dayGan = calendar.GanZhiDDString.Substring(0, 1);
                dayZhi = calendar.GanZhiDDString.Substring(1, 1);
            }

            hourGan = calendar.ChineseHour.Substring(0, 1);
            hourZhi = calendar.ChineseHour.Substring(1, 1);

            System.DateTime lichunDate = System.DateTime.Parse(calendar.ChineseTwentyFour[1]);
            if (context.Result.SolarBirthDate < lichunDate)
            {
                var prevYearCalendar = new Ecan.EcanChineseCalendar(new System.DateTime(birthDate.Year - 1, 12, 31));
                yearGan = prevYearCalendar.GanZhiYYString.Substring(0, 1);
                yearZhi = prevYearCalendar.GanZhiYYString.Substring(1, 1);
            }

            context.CUE1 = AstrologyHelper.SkyToNumber(yearGan); context.CUF1 = AstrologyHelper.FloorToNumber(yearZhi); context.CUE2 = AstrologyHelper.SkyToNumber(monthGan); context.CUF2 = AstrologyHelper.FloorToNumber(monthZhi); context.CUE3 = AstrologyHelper.SkyToNumber(dayGan); context.CUF3 = AstrologyHelper.FloorToNumber(dayZhi); context.CUE4 = AstrologyHelper.SkyToNumber(hourGan); context.CUF4 = AstrologyHelper.FloorToNumber(hourZhi);

            string dayMaster = dayGan;
            var ganStrings = new[] { yearGan, monthGan, dayGan, hourGan };
            var zhiStrings = new[] { yearZhi, monthZhi, dayZhi, hourZhi };
            var zhiInts = new[] { context.CUF1, context.CUF2, context.CUF3, context.CUF4 };
            var pillars = new PillarInfo[4];
            for (int i = 0; i < 4; i++)
            {
                string ganZhiPair = ganStrings[i] + zhiStrings[i];
                string naYin = NaYinMap.GetValueOrDefault(ganZhiPair, "未知");
                string liuShen = (i == 2) ? "" : GetLiuShen(dayMaster, ganStrings[i]);
                var hiddenStemLiuShenList = new List<string>();
                if (HiddenStemsMap.TryGetValue(zhiInts[i], out var hiddenStems))
                {
                    foreach (var stem in hiddenStems)
                    {
                        hiddenStemLiuShenList.Add(GetLiuShen(dayMaster, stem)); hiddenStemLiuShenList.Add(stem);
                    }
                }
                pillars[i] = new PillarInfo(ganStrings[i], zhiStrings[i], liuShen, naYin, hiddenStemLiuShenList);
            }
            var baziInfo = new BaziInfo(pillars[0], pillars[1], pillars[2], pillars[3], calendar.AnimalString, dayMaster);
            context.Result = context.Result with { Bazi = baziInfo };
        }

        private void DetermineBaziLuckCycles(AstrologyCalculationContext context)
        {
            var luckCycles = new List<BaziLuckCycle>();
            var calendar = context.Calendar;
            DateTime birthTime = context.Result.SolarBirthDate;

            bool isYearGanYang = context.CUE1 % 2 != 0;
            bool isMale = context.Request.Gender == 1;
            bool isForward = (isMale && isYearGanYang) || (!isMale && !isYearGanYang);

            DateTime prevJie = DateTime.MinValue;
            DateTime nextJie = DateTime.MaxValue;

            DateTime[] solarTermsDate = new DateTime[calendar.ChineseTwentyFour.Length];
            for (int i = 0; i < calendar.ChineseTwentyFour.Length; i++)
            {
                DateTime.TryParse(calendar.ChineseTwentyFour[i], out solarTermsDate[i]);
            }

            for (int i = 1; i < solarTermsDate.Length; i += 2)
            {
                if (solarTermsDate[i] > birthTime)
                {
                    nextJie = solarTermsDate[i];
                    prevJie = (i > 1) ? solarTermsDate[i - 2] : DateTime.Parse(new EcanChineseCalendar(birthTime.AddMonths(-2)).ChineseTwentyFour[23]);
                    break;
                }
                if (i == 23 && nextJie == DateTime.MaxValue)
                {
                    prevJie = solarTermsDate[23];
                    nextJie = DateTime.Parse(new EcanChineseCalendar(birthTime.AddYears(1)).ChineseTwentyFour[1]);
                }
            }

            TimeSpan diff = isForward ? (nextJie - birthTime) : (birthTime - prevJie);

            double startAgeFloat = diff.TotalDays / 3;
            int startAge = (int)Math.Round(startAgeFloat, MidpointRounding.AwayFromZero);

            int currentGanIndex = context.CUE2;
            int currentZhiIndex = context.CUF2;

            for (int i = 0; i < 8; i++)
            {
                // 【關鍵修正】: 在迴圈的一開始就進行干支的遞增或遞減
                if (isForward)
                {
                    currentGanIndex++;
                    currentZhiIndex++;
                }
                else
                {
                    currentGanIndex--;
                    currentZhiIndex--;
                }

                if (currentGanIndex > 10) currentGanIndex = 1;
                if (currentGanIndex < 1) currentGanIndex = 10;
                if (currentZhiIndex > 12) currentZhiIndex = 1;
                if (currentZhiIndex < 1) currentZhiIndex = 12;

                string gan = S_SKY[currentGanIndex];
                string zhi = S_FLOOR[currentZhiIndex];

                luckCycles.Add(new BaziLuckCycle(
                    startAge + (i * 10),
                    startAge + (i * 10) + 9,
                    gan,
                    zhi,
                    GetLiuShen(context.Result.Bazi.DayMaster, gan)
                ));
            }

            context.Result = context.Result with { BaziLuckCycles = luckCycles };
        }

        private void RunZiWeiChartCalculations(AstrologyCalculationContext context)
        {
            int startGan = 0; switch (context.CUE1) { case 1: case 6: startGan = 3; break; case 2: case 7: startGan = 5; break; case 3: case 8: startGan = 7; break; case 4: case 9: startGan = 9; break; case 5: case 10: startGan = 1; break; }
            for (int i = 0; i < 12; i++) { int palaceIndex = i + 3; if (palaceIndex > 12) palaceIndex -= 12; int currentGan = startGan + i; if (currentGan > 10) currentGan -= 10; context.CCO[palaceIndex] = AstrologyConstants.S_SKY[currentGan]; }
            int month = context.Calendar.ChineseMonth; int hour = AstrologyHelper.GetChineseHourValue(context.Calendar.ChineseHour); int mingGongIndex = 3 + (month - 1) - (hour - 1); while (mingGongIndex < 1) mingGongIndex += 12; context.MingGongIndex = mingGongIndex; int shenGongIndex = 3 + (month - 1) + (hour - 1); while (shenGongIndex > 12) shenGongIndex -= 12; context.ShenGongIndex = shenGongIndex; string mingGongGan = context.CCO[context.MingGongIndex]; string naYin = AstrologyConstants.NaYinSound[(AstrologyHelper.SkyToNumber(mingGongGan), context.MingGongIndex)];
            string wuXingJuText = "";
            switch (naYin.Substring(naYin.Length - 1, 1)) { case "水": context.WuXingJu = 2; wuXingJuText = "水二局"; break; case "木": context.WuXingJu = 3; wuXingJuText = "木三局"; break; case "金": context.WuXingJu = 4; wuXingJuText = "金四局"; break; case "土": context.WuXingJu = 5; wuXingJuText = "土五局"; break; case "火": context.WuXingJu = 6; wuXingJuText = "火六局"; break; }
            context.WuXingJuText = wuXingJuText;
            context.CCB[context.MingGongIndex] = "命宮"; string[] palaceNames = { "兄弟", "夫妻", "子女", "財帛", "疾厄", "遷移", "奴僕", "官祿", "田宅", "福德", "父母" }; int currentPalaceIndex = context.MingGongIndex; foreach (var name in palaceNames) { currentPalaceIndex--; if (currentPalaceIndex < 1) currentPalaceIndex = 12; context.CCB[currentPalaceIndex] = name; }
            if (context.MingGongIndex != context.ShenGongIndex) { context.CCB[context.ShenGongIndex] += "身"; }
            for (int i = 1; i <= 12; i++) { context.PalaceShortNames[i] = context.CCB[i].Substring(0, 1); }
            var branchDirections = new Dictionary<int, string> { { 1, "子 北  方" }, { 2, "丑 北東北" }, { 3, "寅 東北方" }, { 4, "卯 東  方" }, { 5, "辰 東南東" }, { 6, "巳 東南方" }, { 7, "午 南  方" }, { 8, "未 南西南" }, { 9, "申 西南方" }, { 10, "酉 西  方" }, { 11, "戌 西北西" }, { 12, "亥 西北方" } }; var palaces = new List<ZiWeiPalace>(); for (int i = 1; i <= 12; i++) { palaces.Add(new ZiWeiPalace(i, context.CCB[i], context.CCO[i], branchDirections[i], new List<string>(), new List<string>(), new List<string>(), "", "", "", "", new List<string>(), new List<string>(), new List<string>())); }
            palaces.Sort((a, b) => a.Index.CompareTo(b.Index)); context.Result = context.Result with { palaces = palaces };
        }
        private void DetermineMajorStars(AstrologyCalculationContext context)
        {
            int day = context.LunarDay; int wuXingJu = context.WuXingJu; string[] mstarTable = AstrologyConstants.MSTAR[wuXingJu].Split(new[] { ',', ' ' }, System.StringSplitOptions.RemoveEmptyEntries); int nor = int.Parse(mstarTable[day - 1]); string[] norTable = AstrologyConstants.CRC[2, nor].Split(','); context.CCM[int.Parse(norTable[0])] = "紫"; context.CCM[int.Parse(norTable[1])] = "機"; context.CCM[int.Parse(norTable[3])] = "陽"; context.CCM[int.Parse(norTable[4])] = "武"; context.CCM[int.Parse(norTable[5])] = "同"; context.CCM[int.Parse(norTable[8])] = "廉"; int sou = nor switch { 1 => 5, 2 => 4, 3 => 3, 4 => 2, 5 => 1, 6 => 12, 7 => 11, 8 => 10, 9 => 9, 10 => 8, 11 => 7, 12 => 6, _ => 0 }; string[] souTable = AstrologyConstants.CRC[1, sou].Split(','); context.CCN[int.Parse(souTable[0])] = "府"; context.CCN[int.Parse(souTable[1])] = "陰"; context.CCN[int.Parse(souTable[2])] = "貪"; context.CCN[int.Parse(souTable[3])] = "巨"; context.CCN[int.Parse(souTable[4])] = "相"; context.CCN[int.Parse(souTable[5])] = "梁"; context.CCN[int.Parse(souTable[6])] = "殺"; context.CCN[int.Parse(souTable[10])] = "破";
        }
        private int PalaceWrap(int val) { while (val > 12) val -= 12; while (val < 1) val += 12; return val; }

        // ==========================================================
        // 【新增方法】: 根據生年地支計算天馬位置
        // 口訣：寅午戌年馬在申, 申子辰年馬在寅, 巳酉丑年馬在亥, 亥卯未年馬在巳
        // ==========================================================
        private int PlaceTienMa(int yearZhi)
        {
            if (yearZhi == 3 || yearZhi == 7 || yearZhi == 11) return 9; // 寅午戌 (3, 7, 11) -> 申 (9)
            if (yearZhi == 9 || yearZhi == 1 || yearZhi == 5) return 3; // 申子辰 (9, 1, 5) -> 寅 (3)
            if (yearZhi == 6 || yearZhi == 10 || yearZhi == 2) return 12; // 巳酉丑 (6, 10, 2) -> 亥 (12)
            return 6; // 亥卯未 (12, 4, 8) -> 巳 (6)
        }

        // ==========================================================
        // 【新增方法】: 根據生月計算天月位置
        // 口訣：一犬二蛇三在龍四虎五羊六兔宮七豬八羊九在虎十馬冬犬臘寅中
        // (此處採用較常見的天月口訣，需與您的資料核對)
        // ==========================================================
        private int PlaceTienYue(int month)
        {
            // 月份 (1-12) 對應 地支 (1-12)
            int[] posMap = { 11, 6, 5, 3, 8, 4, 12, 8, 3, 7, 11, 3 }; // 戌, 巳, 辰, 寅, 未, 卯, 亥, 未, 寅, 午, 戌, 寅
            if (month >= 1 && month <= 12)
            {
                return posMap[month - 1];
            }
            return 0; // 錯誤處理
        }

        // ==========================================================
        // 【新增方法】: 根據生年天干計算截路空亡 (用於旬空)
        // 口訣：甲己之年申酉, 乙庚之年午未, 丙辛之年辰巳, 丁壬之年寅卯, 戊癸之年子丑
        // ==========================================================
        private (int Pos1, int Pos2) PlaceJieLuKongWang(int yearGan)
        {
            switch (yearGan)
            {
                case 1: // 甲
                case 6: // 己
                    return (9, 10); // 申(9), 酉(10)
                case 2: // 乙
                case 7: // 庚
                    return (7, 8); // 午(7), 未(8)
                case 3: // 丙
                case 8: // 辛
                    return (5, 6); // 辰(5), 巳(6)
                case 4: // 丁
                case 9: // 壬
                    return (3, 4); // 寅(3), 卯(4)
                case 5: // 戊
                case 10: // 癸
                    return (1, 2); // 子(1), 丑(2)
                default:
                    return (0, 0);
            }
        }

        // ==========================================================
        // 【新增方法】: 根據生年地支計算破碎位置
        // 口訣：子午卯酉巳, 寅申巳亥雞(酉), 辰戌丑未丑
        // ==========================================================
        private int PlacePoSui(int yearZhi)
        {
            if (yearZhi == 1 || yearZhi == 7 || yearZhi == 4 || yearZhi == 10) return 6; // 子午卯酉 -> 巳(6)
            if (yearZhi == 3 || yearZhi == 9 || yearZhi == 6 || yearZhi == 12) return 10; // 寅申巳亥 -> 酉(10)
            return 2; // 辰戌丑未 (5, 11, 2, 8) -> 丑(2)
        }

        // ==========================================================
        // 【新增方法】: 根據生月計算陰煞位置
        // 口訣：正七月在寅, 二八月在子, 三九月在戌, 四十月在申, 五十一在午, 六十二在辰
        // ==========================================================
        private int PlaceYinSha(int month)
        {
            switch (month)
            {
                case 1: case 7: return 3; // 寅(3)
                case 2: case 8: return 1; // 子(1)
                case 3: case 9: return 11; // 戌(11)
                case 4: case 10: return 9; // 申(9)
                case 5: case 11: return 7; // 午(7)
                case 6: case 12: return 5; // 辰(5)
                default: return 0;
            }
        }

        // ==========================================================
        // 【新增方法】: 根據生月計算解神位置
        // 口訣：正二在申, 三四在戌, 五六在子, 七八在寅, 九十月坐於辰宮, 十一十二在午宮
        // ==========================================================
        private int PlaceJieShen(int month)
        {
            switch (month)
            {
                case 1: case 2: return 9; // 申(9)
                case 3: case 4: return 11; // 戌(11)
                case 5: case 6: return 1; // 子(1)
                case 7: case 8: return 3; // 寅(3)
                case 9: case 10: return 5; // 辰(5)
                case 11: case 12: return 7; // 午(7)
                default: return 0;
            }
        }
        private void DetermineAllAuxiliaryAndMinorStars(AstrologyCalculationContext context)
        {
            // 獲取生時地支數 (子=1, 丑=2, ...)
            int hourBranchNum = context.CUF4;
            // 獲取生月地支數 (寅=3, 卯=4, ...)
            int monthBranchNum = context.CUF2;
            // 獲取生年天干數 (甲=1, 乙=2, ...)
            int yearStemNum = context.CUE1;
            int yearGan = context.CUE1; int yearZhi = context.CUF1; int month = context.Calendar.ChineseMonth;
            int day = context.LunarDay;
            int hour = AstrologyHelper.GetChineseHourValue(context.Calendar.ChineseHour); int hourZhi = context.CUF4;
            context.ZuoFuPos = PalaceWrap(5 + month - 1); context.YouBiPos = PalaceWrap(11 - (month - 1));
            context.WenChangPos = PalaceWrap(11 - (hour - 1)); context.WenQuPos = PalaceWrap(5 + hour - 1);
            int IwPos = PalaceWrap(2 + month - 1);  
            int SiPos = PalaceWrap(10 + month - 1); 
            context.SecondaryStars[context.ZuoFuPos] += "左 "; context.SecondaryStars[context.YouBiPos] += "右 ";
            context.SecondaryStars[context.WenChangPos] += "昌 "; context.SecondaryStars[context.WenQuPos] += "曲 ";
            context.SecondaryStars[IwPos] += "姚 "; context.SecondaryStars[SiPos] += "刑 ";
            int WenChangPos = context.WenChangPos;
            int WenQuPos = context.WenQuPos;
            // context.GoodStars[PalaceWrap(context.WenChangPos + day - 2)] += "恩 "; context.GoodStars[PalaceWrap(context.WenQuPos + day - 2)] += "貴 ";
            if (yearGan % 2 != 0) { context.SecondaryStars[PalaceWrap(2 + (yearGan + 1) / 2 - 1)] += "魁 "; context.SecondaryStars[PalaceWrap(8 - ((yearGan + 1) / 2 - 1))] += "鉞 "; } else { context.SecondaryStars[PalaceWrap(9 - (yearGan / 2 - 1))] += "魁 "; context.SecondaryStars[PalaceWrap(3 + (yearGan / 2 - 1))] += "鉞 "; }
            int luCunPos = 0; switch (yearGan) { case 1: luCunPos = 3; break; case 2: luCunPos = 4; break; case 3: luCunPos = 6; break; case 4: luCunPos = 7; break; case 5: luCunPos = 6; break; case 6: luCunPos = 7; break; case 7: luCunPos = 9; break; case 8: luCunPos = 10; break; case 9: luCunPos = 12; break; case 10: luCunPos = 1; break; }
            context.LuCunPos = luCunPos;
            context.SecondaryStars[luCunPos] += "祿 "; context.BadStars[PalaceWrap(luCunPos + 1)] += "羊 "; context.BadStars[PalaceWrap(luCunPos - 1)] += "陀 ";
            int huoPosBase = new int[] { 2, 3, 4, 10, 2, 3, 4, 10, 2, 3, 4, 10 }[yearZhi - 1];
            int lingPosBase = new int[] { 4, 11, 11, 11, 4, 11, 11, 11, 4, 11, 11, 11 }[yearZhi - 1];
            context.BadStars[PalaceWrap(huoPosBase + hour - 1)] += "火 "; context.BadStars[PalaceWrap(lingPosBase - (hour - 1))] += "鈴 ";
            context.BadStars[hourZhi] += "劫 "; context.BadStars[PalaceWrap(2 + 12 - hourZhi)] += "空 ";
            System.Action<string, string> AddFourTransformation = (starName, transType) => { for (int i = 1; i <= 12; i++) { string allStarsInPalace = context.CCM[i] + context.CCN[i] + context.SecondaryStars[i]; if (allStarsInPalace.Contains(starName)) { context.FourTransformationStars[i] += transType; return; } } };
            switch (yearGan) { case 1: AddFourTransformation("廉", "祿"); AddFourTransformation("破", "權"); AddFourTransformation("武", "科"); AddFourTransformation("陽", "忌"); break; case 2: AddFourTransformation("機", "祿"); AddFourTransformation("梁", "權"); AddFourTransformation("紫", "科"); AddFourTransformation("陰", "忌"); break; case 3: AddFourTransformation("同", "祿"); AddFourTransformation("機", "權"); AddFourTransformation("昌", "科"); AddFourTransformation("廉", "忌"); break; case 4: AddFourTransformation("陰", "祿"); AddFourTransformation("同", "權"); AddFourTransformation("機", "科"); AddFourTransformation("巨", "忌"); break; case 5: AddFourTransformation("貪", "祿"); AddFourTransformation("陰", "權"); AddFourTransformation("右", "科"); AddFourTransformation("機", "忌"); break; case 6: AddFourTransformation("武", "祿"); AddFourTransformation("貪", "權"); AddFourTransformation("梁", "科"); AddFourTransformation("曲", "忌"); break; case 7: AddFourTransformation("陽", "祿"); AddFourTransformation("武", "權"); AddFourTransformation("陰", "科"); AddFourTransformation("同", "忌"); break; case 8: AddFourTransformation("巨", "祿"); AddFourTransformation("陽", "權"); AddFourTransformation("曲", "科"); AddFourTransformation("昌", "忌"); break; case 9: AddFourTransformation("梁", "祿"); AddFourTransformation("紫", "權"); AddFourTransformation("左", "科"); AddFourTransformation("武", "忌"); break; case 10: AddFourTransformation("破", "祿"); AddFourTransformation("巨", "權"); AddFourTransformation("陰", "科"); AddFourTransformation("貪", "忌"); break; }

            string[] doctorStars = { "博", "力", "青", "小", "將", "奏", "蜚", "喜", "病", "大", "伏", "官" };
            string[] ageStars = { "歲", "晦", "喪", "貫", "官", "小", "大", "龍", "虎", "德", "弔", "病" };
            string[] generalStars = { "將", "鞍", "驛", "息", "蓋", "劫", "災", "天", "指", "咸", "月", "亡" };
            bool isForward = (context.Request.Gender == 1 && (yearGan % 2 != 0)) || (context.Request.Gender == 2 && (yearGan % 2 == 0));
            var doctorStarPositions = new Dictionary<int, string>(); for (int i = 0; i < 12; i++) { int pos = isForward ? PalaceWrap(context.LuCunPos + i) : PalaceWrap(context.LuCunPos - i); doctorStarPositions[pos] = doctorStars[i]; }
            var ageStarPositions = new Dictionary<int, string>(); for (int i = 0; i < 12; i++) { int pos = PalaceWrap(yearZhi + i); ageStarPositions[pos] = ageStars[i]; }
            var generalStarPositions = new Dictionary<int, string>(); int generalStartPos = new int[] { 7, 8, 9, 4, 5, 6, 1, 2, 3, 10, 11, 12 }[yearZhi - 1]; for (int i = 0; i < 12; i++) { int pos = PalaceWrap(generalStartPos + i); generalStarPositions[pos] = generalStars[i]; }

            for (int i = 1; i <= 12; i++)
            {
                var stars = new List<string> {
                    doctorStarPositions.GetValueOrDefault(i, ""),
                    ageStarPositions.GetValueOrDefault(i, ""),
                    generalStarPositions.GetValueOrDefault(i, "")
                };
                context.SmallStars[i] = string.Join("|", stars);
            }

            context.BadStars[PalaceWrap(5 - yearZhi)] += "鸞 "; context.GoodStars[PalaceWrap(11 - yearZhi)] += "喜 ";
            context.BadStars[PalaceWrap(7 + yearZhi - 1)] += "哭 "; context.BadStars[PalaceWrap(7 - (yearZhi - 1))] += "虛 ";
            context.GoodStars[PalaceWrap(5 + yearZhi - 1)] += "龍 "; context.GoodStars[PalaceWrap(11 - (yearZhi - 1))] += "鳳 ";
            int[] guPosMap = { 3, 3, 3, 6, 6, 6, 9, 9, 9, 12, 12, 12 }; int[] guaPosMap = { 11, 11, 11, 2, 2, 2, 5, 5, 5, 8, 8, 8 };
            context.BadStars[guPosMap[yearZhi - 1]] += "孤 "; context.BadStars[guaPosMap[yearZhi - 1]] += "寡 ";
            context.GoodStars[PalaceWrap(context.ZuoFuPos + day - 1)] += "三 "; context.GoodStars[PalaceWrap(context.YouBiPos - day + 1)] += "八 ";
            //少判斷文昌文曲在命宮或身宮時，不加恩貴
            context.GoodStars[PalaceWrap(WenChangPos + day - 2)] += "恩 "; // context.GoodStars[PalaceWrap(WenQuPos + day)] += "貴 ";
            if (WenQuPos != -1)
            {
                // 1. 從文曲宮(wenQuPos)起算初一
                // 2. 順數到生日 (lunarDay)
                // 3. 退一步 (逆數一宮)
                int stepForward = day - 1; // 從初一(0步)到生日(lunarDay-1步)
                int tianGuiPos = (WenQuPos + stepForward - 1) % 12 + 1; // 順數到生日所在宮位

                // 再退一步 (逆時針一宮)
                tianGuiPos = tianGuiPos + 1;
                if (tianGuiPos == 0) tianGuiPos = 12;
                context.GoodStars[tianGuiPos] += (string.IsNullOrEmpty(context.GoodStars[tianGuiPos]) ? "" : "貴") ;
            }
            // context.GoodStars[PalaceWrap(context.WenChangPos + day - 2)] += "恩 "; context.GoodStars[PalaceWrap(context.WenQuPos + day - 2)] += "貴 ";
            context.GoodStars[PalaceWrap(context.MingGongIndex + day - 1)] += "才 ";
            context.GoodStars[PalaceWrap(context.ShenGongIndex + day - 1)] += "壽 ";

            // --- 【修正】地劫、天空 (論生時) ---
            // 口訣：亥上起子順安劫，逆去便是地空鄉。

            // 地劫：起亥宮(12)順數到生時
            int diJiePos = (12 + hourBranchNum - 1) % 12 ; // 12 + hour - 1 (起數在亥=12), % 12 + 1
                                                              // 天空：起亥宮(12)逆數到生時
            int tianKongPos = 12 - (hourBranchNum - 1); // 12 - (hour - 1)

            context.SecondaryStars[diJiePos] += (string.IsNullOrEmpty(context.SecondaryStars[diJiePos]) ? "" : " ") + "劫";
            context.SecondaryStars[tianKongPos] += (string.IsNullOrEmpty(context.SecondaryStars[tianKongPos]) ? "" : " ") + "空";

            // 獲取農曆生日
            int lunarDay = context.LunarDay;

            //台輔封誥
            // --- 【修正】台輔、封誥 (論生時) ---
            // 口訣：台輔星從午宮起子，順至生時是貴鄉。/ 封誥寅宮來起子，順到生時是貴方。

            // 台輔：起午宮(7)順數到生時
            int taiFuPos = (7 + hourBranchNum - 1) % 12;//+ 1; // 7(午) + hour - 1 (起數在午=7)
                                                        // 封誥：起寅宮(3)順數到生時
            int fengGaoPos = (3 + hourBranchNum - 1) % 12;// + 1; // 3(寅) + hour - 1 (起數在寅=3)

            context.SmallStars[taiFuPos] += (string.IsNullOrEmpty(context.SmallStars[taiFuPos]) ? "" : "|") + "輔";
            context.SmallStars[fengGaoPos] += (string.IsNullOrEmpty(context.SmallStars[fengGaoPos]) ? "" : "|") + "誥";

            // ==========================================================
            // 【新增星曜的安星邏輯】
            // ==========================================================

            // 1. 天馬 (論年支)
            int tianMaPos = PlaceTienMa(yearZhi);
            context.SmallStars[tianMaPos] = (string.IsNullOrEmpty(context.SmallStars[tianMaPos]) ? "" : context.SmallStars[tianMaPos] + "|") + "馬";

            // 2. 天月 (論生月)
            int tianYuePos = PlaceTienYue(month);
            context.SmallStars[tianYuePos] = (string.IsNullOrEmpty(context.SmallStars[tianYuePos]) ? "" : context.SmallStars[tianYuePos] + "|") + "月";

            // 3. 天哭、天虛 (論年支) (保留原邏輯並加入SmallStars)
            // 原有邏輯：
            // context.BadStars[PalaceWrap(7 + yearZhi - 1)] += "哭 "; context.BadStars[PalaceWrap(7 - (yearZhi - 1))] += "虛 ";
            // 修正為：
            int tianKuPos = PalaceWrap(7 + yearZhi - 1); // 午(7)起子年，順行年支 (天哭逆行) -> 應為 7 - (yearZhi - 1)
            int tianXuPos = PalaceWrap(7 - (yearZhi - 1)); // 午(7)起子年，逆行年支 (天虛順轉) -> 應為 7 + (yearZhi - 1)

            // 根據「哭逆行兮虛順轉，數到生年便停留」的常見理解：
            // 午宮起子年（子=1），哭逆轉，虛順轉
            tianKuPos = PalaceWrap(7 - (yearZhi - 1));
            tianXuPos = PalaceWrap(7 + (yearZhi - 1));

            context.SmallStars[tianKuPos] = (string.IsNullOrEmpty(context.SmallStars[tianKuPos]) ? "" : context.SmallStars[tianKuPos] + "|") + "哭";
            context.SmallStars[tianXuPos] = (string.IsNullOrEmpty(context.SmallStars[tianXuPos]) ? "" : context.SmallStars[tianXuPos] + "|") + "虛";

            // 4. 旬空 (以截路空亡處理，論年干)
            var (jieLu1, jieLu2) = PlaceJieLuKongWang(yearGan);
            if (jieLu1 != 0) context.SmallStars[jieLu1] = (string.IsNullOrEmpty(context.SmallStars[jieLu1]) ? "" : context.SmallStars[jieLu1] + "|") + "截";
            if (jieLu2 != 0) context.SmallStars[jieLu2] = (string.IsNullOrEmpty(context.SmallStars[jieLu2]) ? "" : context.SmallStars[jieLu2] + "|") + "路";

            // 5. 破碎 (論年支)
            int poSuiPos = PlacePoSui(yearZhi);
            context.SmallStars[poSuiPos] = (string.IsNullOrEmpty(context.SmallStars[poSuiPos]) ? "" : context.SmallStars[poSuiPos] + "|") + "碎";

            // 6. 天壽 (論身宮，年支) (保留原邏輯並加入SmallStars)
            // 原有邏輯： context.GoodStars[PalaceWrap(context.ShenGongIndex + day - 1)] += "壽 ";
            // 修正為：
            int tianShouPos = PalaceWrap(context.ShenGongIndex + yearZhi - 1); // 身宮起子年，順行年支
            context.SmallStars[tianShouPos] = (string.IsNullOrEmpty(context.SmallStars[tianShouPos]) ? "" : context.SmallStars[tianShouPos] + "|") + "壽";

            // 7. 天才 (論命宮，年支) (保留原邏輯並加入SmallStars)
            // 原有邏輯： context.GoodStars[PalaceWrap(context.MingGongIndex + day - 1)] += "才 ";
            // 修正為：
            int tianCaiPos = PalaceWrap(context.MingGongIndex + yearZhi - 1); // 命宮起子年，順行年支
            context.SmallStars[tianCaiPos] = (string.IsNullOrEmpty(context.SmallStars[tianCaiPos]) ? "" : context.SmallStars[tianCaiPos] + "|") + "才";

            // 8. 陰煞 (論生月)
            int yinShaPos = PlaceYinSha(month);
            context.SmallStars[yinShaPos] = (string.IsNullOrEmpty(context.SmallStars[yinShaPos]) ? "" : context.SmallStars[yinShaPos] + "|") + "陰";

            // 9. 解神 (論生月)
            int jieShenPos = PlaceJieShen(month);
            context.SmallStars[jieShenPos] = (string.IsNullOrEmpty(context.SmallStars[jieShenPos]) ? "" : context.SmallStars[jieShenPos] + "|") + "解";
        }
        private void DetermineDecadeLuckCycles(AstrologyCalculationContext context)
        {
            int startAge = context.WuXingJu;
            bool isForward = (context.Request.Gender == 1 && context.CUE1 % 2 != 0) || (context.Request.Gender == 2 && context.CUE1 % 2 == 0);
            for (int i = 0; i < 12; i++) { int palaceIndex = isForward ? PalaceWrap(context.MingGongIndex + i) : PalaceWrap(context.MingGongIndex - i); context.CCX[palaceIndex] = $"{startAge}-{startAge + 9}"; startAge += 10; }
        }
        private void DetermineLifeCycleStars(AstrologyCalculationContext context)
        {
            string[] lifeCycleNames = { "長", "沐", "冠", "臨", "帝", "衰", "病", "死", "墓", "絕", "胎", "養" };
            bool isForward = (context.Request.Gender == 1 && (context.CUE1 % 2 != 0)) || (context.Request.Gender == 2 && (context.CUE1 % 2 == 0));
            var startPosMap = new Dictionary<int, int> { { 2, 9 }, { 3, 12 }, { 4, 6 }, { 5, 9 }, { 6, 3 } };
            if (!startPosMap.TryGetValue(context.WuXingJu, out int startPos)) { startPos = 3; }
            for (int i = 0; i < 12; i++) { int pos = isForward ? PalaceWrap(startPos + i) : PalaceWrap(startPos - i); context.LifeCycleStage[pos] = lifeCycleNames[i]; }
        }
        private void DetermineStarBrightness(AstrologyCalculationContext context)
        {
            for (int i = 1; i <= 12; i++)
            {
                int semanticIndex = N12S.IndexOf(context.PalaceShortNames[i]) + 1;
                if (semanticIndex <= 0 || semanticIndex > 12) continue;
                var brightnessParts = new List<string>();
                if (!string.IsNullOrEmpty(context.CCM[i])) { int starIndexJ = M6S.IndexOf(context.CCM[i]) + 1; if (starIndexJ > 0) { string numericCode = NUMERIC_LIGHT_TABLE[semanticIndex].Substring(starIndexJ - 1, 1); string symbolCode = SYMBOL_LIGHT_TABLE[i].Substring(starIndexJ - 1, 1); brightnessParts.Add($"{numericCode}{symbolCode}"); } }
                if (!string.IsNullOrEmpty(context.CCN[i])) { int starIndexJ = N8S.IndexOf(context.CCN[i]) + 1; if (starIndexJ > 0) { string numericCode = NUMERIC_LIGHT_TABLE[semanticIndex].Substring(starIndexJ + 5, 1); string symbolCode = SYMBOL_LIGHT_TABLE[i].Substring(starIndexJ + 5, 1); brightnessParts.Add($"{numericCode}{symbolCode}"); } }
                context.MainStarBrightness[i] = string.Join(",", brightnessParts);
            }
        }

        private string CalculatePalaceStemTransformations(AstrologyCalculationContext context, string palaceStem)
        {
            var starMap = new Dictionary<string, string[]> { { "甲", new[] { "廉", "破", "武", "陽" } }, { "乙", new[] { "機", "梁", "紫", "陰" } }, { "丙", new[] { "同", "機", "昌", "廉" } }, { "丁", new[] { "陰", "同", "機", "巨" } }, { "戊", new[] { "貪", "陰", "右", "機" } }, { "己", new[] { "武", "貪", "梁", "曲" } }, { "庚", new[] { "陽", "武", "陰", "同" } }, { "辛", new[] { "巨", "陽", "曲", "昌" } }, { "壬", new[] { "梁", "紫", "左", "武" } }, { "癸", new[] { "破", "巨", "陰", "貪" } } };
            if (!starMap.ContainsKey(palaceStem)) return "";
            string[] targetStars = starMap[palaceStem];
            var result = new StringBuilder();
            foreach (var star in targetStars) { bool found = false; for (int i = 1; i <= 12; i++) { string allStarsInPalace = context.CCM[i] + context.CCN[i] + context.SecondaryStars[i]; if (allStarsInPalace.Contains(star)) { result.Append(context.PalaceShortNames[i]); found = true; break; } } if (!found) result.Append(" "); }
            return result.ToString();
        }
        private void UpdateFinalResult(AstrologyCalculationContext context)
        {
            string[] mingZhuMap = { "貪", "巨", "祿", "文", "廉", "武", "破", "武", "廉", "文", "祿", "巨" };
            string[] shenZhuMap = { "火", "相", "梁", "同", "昌", "機", "火", "相", "梁", "同", "昌", "機" };
            context.MingZhu = mingZhuMap[context.MingGongIndex - 1];
            context.ShenZhu = shenZhuMap[context.CUF1 - 1];
            var finalPalaces = new List<ZiWeiPalace>();
            for (int i = 1; i <= 12; i++)
            {
                var palace = context.Result.palaces.First(p => p.Index == i);
                var majorStars = new List<string>();
                if (!string.IsNullOrEmpty(context.CCM[i])) majorStars.Add(context.CCM[i].Trim());
                if (!string.IsNullOrEmpty(context.CCN[i])) majorStars.Add(context.CCN[i].Trim());
                var secondaryStars = context.SecondaryStars[i]?.Split(new[] { ' ' }, System.StringSplitOptions.RemoveEmptyEntries).ToList() ?? new List<string>();
                var goodStars = context.GoodStars[i]?.Split(new[] { ' ' }, System.StringSplitOptions.RemoveEmptyEntries).ToList() ?? new List<string>();
                //var badStars = context.BadStars[i]?.Split(new[] { ' ' }, System.StringSplitOptions.RemoveEmptyEntries).ToList() ?? new List<string>();
                var badStars = context.BadStars[i]?.Split(new[] { ' ' }, System.StringSplitOptions.RemoveEmptyEntries).ToList() ?? new List<string>();
                var smallStars = context.SmallStars[i]?.Split(new[] { '|' }, System.StringSplitOptions.RemoveEmptyEntries).ToList() ?? new List<string>();
                //var smallStars = context.SmallStars[i]?.Split(new[] { '|' }, System.StringSplitOptions.None).ToList() ?? new List<string>();
                var annualStarTransformations = context.FourTransformationStars[i]?.ToCharArray().Select(c => c.ToString()).ToList() ?? new List<string>();
                string palaceStemTrans = CalculatePalaceStemTransformations(context, context.CCO[i]);
                finalPalaces.Add(palace with { MajorStars = majorStars, SecondaryStars = secondaryStars, AnnualStarTransformations = annualStarTransformations, DecadeAgeRange = context.CCX[i], LifeCycleStage = context.LifeCycleStage[i] ?? "", MainStarBrightness = context.MainStarBrightness[i] ?? "", PalaceStemTransformations = palaceStemTrans, GoodStars = goodStars, BadStars = badStars, SmallStars = smallStars });
            }
            context.Result = context.Result with
            {
                palaces = finalPalaces,
                WuXingJuText = context.WuXingJuText,
                MingZhu = context.MingZhu,
                ShenZhu = context.ShenZhu
            };
        }
    }
}