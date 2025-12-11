using Ecan;
using Ecanapi.Data;
using Ecanapi.Models;
using Ecanapi.Models.Ecanapi.Models;
using Ecanapi.Services.AstrologyEngine;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace Ecanapi.Services
{
    public class AstrologyService : IAstrologyService
    {
        // ⭐【新增】BlindSchoolUltimateAnalyzer 欄位
        private readonly BlindSchoolUltimateAnalyzer _analyzer = new(); // 必須 new 一個實例
        private static readonly Dictionary<string, List<string>> BranchToHiddenStems = new()
        {
            {"子", new List<string> { "癸" }},
            {"丑", new List<string> { "己", "癸", "辛" }},
            {"寅", new List<string> { "甲", "丙", "戊" }},
            {"卯", new List<string> { "乙" }},
            {"辰", new List<string> { "戊", "乙", "癸" }},
            {"巳", new List<string> { "丙", "戊", "庚" }},
            {"午", new List<string> { "丁", "己" }},
            {"未", new List<string> { "己", "乙", "丁" }},
            {"申", new List<string> { "庚", "壬", "戊" }},
            {"酉", new List<string> { "辛" }},
            {"戌", new List<string> { "戊", "辛", "丁" }},
            {"亥", new List<string> { "壬", "甲" }},
        };
        private static readonly Dictionary<string, Dictionary<string, string>> HeavenlyStemLiuShenTable =
            new Dictionary<string, Dictionary<string, string>>
            {
                ["甲"] = new Dictionary<string, string>
                {
                    ["甲"] = "比",
                    ["乙"] = "劫",
                    ["丙"] = "食",
                    ["丁"] = "傷",
                    ["戊"] = "才",
                    ["己"] = "財",
                    ["庚"] = "殺",
                    ["辛"] = "官",
                    ["壬"] = "梟",
                    ["癸"] = "印"
                },
                ["乙"] = new Dictionary<string, string>
                {
                    ["乙"] = "比",
                    ["甲"] = "劫",
                    ["丁"] = "食",
                    ["丙"] = "傷",
                    ["己"] = "才",
                    ["戊"] = "財",
                    ["辛"] = "殺",
                    ["庚"] = "官",
                    ["癸"] = "梟",
                    ["壬"] = "印"
                },
                ["丙"] = new Dictionary<string, string>
                {
                    ["甲"] = "印",
                    ["乙"] = "梟",
                    ["丙"] = "比",
                    ["丁"] = "劫",
                    ["戊"] = "食",
                    ["己"] = "傷",
                    ["庚"] = "才",
                    ["辛"] = "財",
                    ["壬"] = "殺",
                    ["癸"] = "官"
                },
                ["丁"] = new Dictionary<string, string>
                {
                    ["甲"] = "印",
                    ["乙"] = "梟",
                    ["丙"] = "劫",
                    ["丁"] = "比",
                    ["戊"] = "食",
                    ["己"] = "傷",
                    ["庚"] = "才",
                    ["辛"] = "財",
                    ["壬"] = "殺",
                    ["癸"] = "官"
                },
                ["戊"] = new Dictionary<string, string>
                {
                    ["甲"] = "殺",
                    ["乙"] = "官",
                    ["丙"] = "印",
                    ["丁"] = "梟",
                    ["戊"] = "比",
                    ["己"] = "劫",
                    ["庚"] = "食",
                    ["辛"] = "傷",
                    ["壬"] = "才",
                    ["癸"] = "財"
                },
                ["己"] = new Dictionary<string, string>
                {
                    ["甲"] = "殺",
                    ["乙"] = "官",
                    ["丙"] = "印",
                    ["丁"] = "梟",
                    ["戊"] = "劫",
                    ["己"] = "比",
                    ["庚"] = "食",
                    ["辛"] = "傷",
                    ["壬"] = "才",
                    ["癸"] = "財"
                },
                ["庚"] = new Dictionary<string, string>
                {
                    ["甲"] = "才",
                    ["乙"] = "財",
                    ["丙"] = "殺",
                    ["丁"] = "官",
                    ["戊"] = "印",
                    ["己"] = "梟",
                    ["庚"] = "比",
                    ["辛"] = "劫",
                    ["壬"] = "食",
                    ["癸"] = "傷"
                },
                ["辛"] = new Dictionary<string, string>
                {
                    ["甲"] = "傷",
                    ["乙"] = "財",
                    ["丙"] = "官",
                    ["丁"] = "殺",
                    ["戊"] = "印",
                    ["己"] = "梟",
                    ["庚"] = "比",
                    ["辛"] = "劫",
                    ["壬"] = "傷",
                    ["癸"] = "食"
                },
                ["壬"] = new Dictionary<string, string>
                {
                    ["甲"] = "食",
                    ["乙"] = "傷",
                    ["丙"] = "才",
                    ["丁"] = "財",
                    ["戊"] = "殺",
                    ["己"] = "官",
                    ["庚"] = "印",
                    ["辛"] = "梟",
                    ["壬"] = "比",
                    ["癸"] = "劫"
                },
                ["癸"] = new Dictionary<string, string>
                {
                    ["甲"] = "食",
                    ["乙"] = "傷",
                    ["丙"] = "才",
                    ["丁"] = "財",
                    ["戊"] = "殺",
                    ["己"] = "官",
                    ["庚"] = "印",
                    ["辛"] = "梟",
                    ["壬"] = "劫",
                    ["癸"] = "比"
                }
            };

        private readonly ICalendarService _calendarService;

        public AstrologyService(ICalendarService calendarService)
        {
            _calendarService = calendarService;
        }

        #region --- Bazi & Brightness Data ---
        private static readonly Dictionary<string, string> NaYinMap = new Dictionary<string, string> { { "甲子", "海中金" }, { "乙丑", "海中金" }, { "丙寅", "爐中火" }, { "丁卯", "爐中火" }, { "戊辰", "大林木" }, { "己巳", "大林木" }, { "庚午", "路旁土" }, { "辛未", "路旁土" }, { "壬申", "劍鋒金" }, { "癸酉", "劍鋒金" }, { "甲戌", "山頭火" }, { "乙亥", "山頭火" }, { "丙子", "澗下水" }, { "丁丑", "澗下水" }, { "戊寅", "城頭土" }, { "己卯", "城頭土" }, { "庚辰", "白蠟金" }, { "辛巳", "白蠟金" }, { "壬午", "楊柳木" }, { "癸未", "楊柳木" }, { "甲申", "泉中水" }, { "乙酉", "泉中水" }, { "丙戌", "屋上土" }, { "丁亥", "屋上土" }, { "戊子", "霹靂火" }, { "己丑", "霹靂火" }, { "庚寅", "松柏木" }, { "辛卯", "松柏木" }, { "壬辰", "長流水" }, { "癸巳", "長流水" }, { "甲午", "沙中金" }, { "乙未", "沙中金" }, { "丙申", "山下火" }, { "丁酉", "山下火" }, { "戊戌", "平地木" }, { "己亥", "平地木" }, { "庚子", "壁上土" }, { "辛丑", "壁上土" }, { "壬寅", "金箔金" }, { "癸卯", "金箔金" }, { "甲辰", "覆燈火" }, { "乙巳", "覆燈火" }, { "丙午", "天河水" }, { "丁未", "天河水" }, { "戊申", "大驛土" }, { "己酉", "大驛土" }, { "庚戌", "釵釧金" }, { "辛亥", "釵釧金" }, { "壬子", "桑柘木" }, { "癸丑", "桑柘木" }, { "甲寅", "大溪水" }, { "乙卯", "大溪水" }, { "丙辰", "沙中土" }, { "丁巳", "沙中土" }, { "戊午", "天上火" }, { "己未", "天上火" }, { "庚申", "石榴木" }, { "辛酉", "石榴木" }, { "壬戌", "大海水" }, { "癸亥", "大海水" } };
        private static readonly Dictionary<int, string[]> HiddenStemsMap = new Dictionary<int, string[]> { { 1, new[] { "癸" } }, { 2, new[] { "己", "癸", "辛" } }, { 3, new[] { "甲", "丙", "戊" } }, { 4, new[] { "乙" } }, { 5, new[] { "戊", "乙", "癸" } }, { 6, new[] { "丙", "戊", "庚" } }, { 7, new[] { "丁", "己" } }, { 8, new[] { "己", "乙", "丁" } }, { 9, new[] { "庚", "壬", "戊" } }, { 10, new[] { "辛" } }, { 11, new[] { "戊", "辛", "丁" } }, { 12, new[] { "壬", "甲" } } };

        //public string GetLiuShen(string dayStem, string otherStem)
        //{
        //    if (HeavenlyStemLiuShenTable.TryGetValue(dayStem, out var table) && table.TryGetValue(otherStem, out var result))
        //        return result;
        //    return "";
        //}

        private string GetLiuShen(string dayMasterGan, string otherGan) { if (dayMasterGan == otherGan) return "比"; int dayMasterIndex = "甲乙丙丁戊己庚辛壬癸".IndexOf(dayMasterGan); int otherIndex = "甲乙丙丁戊己庚辛壬癸".IndexOf(otherGan); int dayMasterWuXing = dayMasterIndex / 2; int otherWuXing = otherIndex / 2; bool isDayMasterYang = dayMasterIndex % 2 == 0; bool isOtherYang = otherIndex % 2 == 0; if (otherWuXing == (dayMasterWuXing + 1) % 5) return isDayMasterYang == isOtherYang ? "食" : "傷"; if (otherWuXing == (dayMasterWuXing + 2) % 5) return isDayMasterYang == isOtherYang ? "才" : "財"; if (otherWuXing == (dayMasterWuXing + 3) % 5) return isDayMasterYang == isOtherYang ? "殺" : "官"; if (otherWuXing == (dayMasterWuXing + 4) % 5) return isDayMasterYang == isOtherYang ? "梟" : "印"; return "劫"; }
        private static readonly string[] S_SKY = { "", "甲", "乙", "丙", "丁", "戊", "己", "庚", "辛", "壬", "癸" };
        private static readonly string[] S_FLOOR = { "", "子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥" };
        private static readonly string[] NUMERIC_LIGHT_TABLE = { "", "33323133113321", "22313033003300", "23213032003300", "22313032012300", "22313033002201", "32332233122222", "21213032003300", "32332333113313", "21322233103301", "32322322223222", "21232133012221", "21213133003310" };
        private static readonly string[] SYMBOL_LIGHT_TABLE = { "", "+*!+*!+*++!*++", "!!++!+*++!*++!", "+!+++++!!++++-", "!+*++-+!!+!*+-", "++*++++!+!++++", "+!*!+-+!-!+++!", "***+!!+!++!*++", "!!++!+*++!+++!", "++!+!+++!!+!+-", "!+!++-+*!+!!+-", "++!++++*+!++++", "+!!!+-+*-!+++!" };
        private static readonly string N12S = "命父兄夫子財疾遷奴官田福";
        private static readonly string M6S = "紫機陽武同廉";
        private static readonly string N8S = "府陰貪巨相梁殺破";
        #endregion

        // B. ⭐【新增】`AnalyzeAsync` 方法 (從 BaziAnalysisService 複製過來)
        // 由於您想在 AstrologyService 內使用這個功能，但 IAstrologyService 介面中沒有這個簽名。
        // 我將假定您是在 `CalculateChartAsync` 或其他方法中呼叫這個功能。
        // 為了將功能保留在 AstrologyService 中，您可以新增一個私有方法，或將其合併到現有方法中。

        // 如果您希望它能被外部呼叫，請將此方法簽名加入 IAstrologyService 介面中。
        // 這裡我們假設您將它作為一個可以在內部被其他方法呼叫的功能：
        //        private BlindSchoolAnalysisResult AnalyzeBazi(BaziInfo bazi)
        private ShiShenResult AnalyzeBazi(BaziInfo bazi, CsvDataContainer csvData)
        {
            // 1. 將 BaziInfo 中的四柱資訊轉換成干支字串 (Analyzer需要的輸入)
            string yearPillar = bazi.YearPillar.HeavenlyStem + bazi.YearPillar.EarthlyBranch;
            string monthPillar = bazi.MonthPillar.HeavenlyStem + bazi.MonthPillar.EarthlyBranch;
            string dayPillar = bazi.DayPillar.HeavenlyStem + bazi.DayPillar.EarthlyBranch;
            string timePillar = bazi.TimePillar.HeavenlyStem + bazi.TimePillar.EarthlyBranch;       
            char shiShenDayStem = bazi.DayMaster[0]; // 取得日主天干
            // 2. 呼叫核心分析器的同步方法
            var result = _analyzer.AnalyzeShiShen(yearPillar, monthPillar, dayPillar, timePillar,shiShenDayStem, csvData);
            //var result = _analyzer.AnalyzeShiShen(yearPillar, monthPillar, dayPillar, timePillar);
            // 3. 回傳結果
            return result;
        }
        public async Task<AstrologyChartResult> CalculateChartAsync(AstrologyRequest request)
        {
            var context = new AstrologyCalculationContext(request);
            await DetermineCorrectBaziPillars(context);
            // 【新增】在這裡呼叫八字星煞計算
            DetermineBaziShensha(context); 
            DetermineBaziLuckCycles(context);
            // --- 【修正 3：執行盲派八字分析】 ---
            // ⭐ 1. 【執行檔案 I/O】: 在服務層完成所有非同步查詢讀取相關的csv
            var ExpData = await LoadRequiredDataAsync(context.Result.Bazi); // 非同步操作
            var baziAnalysisResult = AnalyzeBazi(context.Result.Bazi, ExpData);
            context.Result = context.Result with { BaziAnalysisResult = baziAnalysisResult };
            if (ExpData.LiuShiJiaZiData != null)
            {
                context.Result.BaziAnalysisResult.ShiShen += "[年柱(根)]:" + ExpData.LiuShiJiaZiData["YearPillar"]["rgxx"]+ ExpData.LiuShiJiaZiData["YearPillar"]["rgcz"];
                context.Result.BaziAnalysisResult.ShiShen += "[月柱(苗)]:" + ExpData.LiuShiJiaZiData["MonthPillar"]["rgxx"] + ExpData.LiuShiJiaZiData["MonthPillar"]["rgcz"];
                context.Result.BaziAnalysisResult.ShiShen += "[日柱(花)]:" + ExpData.LiuShiJiaZiData["DayPillar"]["rgxx"] + ExpData.LiuShiJiaZiData["DayPillar"]["rgcz"];
                context.Result.BaziAnalysisResult.ShiShen += "[時柱(果)]:" + ExpData.LiuShiJiaZiData["TimePillar"]["rgxx"] + ExpData.LiuShiJiaZiData["TimePillar"]["rgcz"];
            }

            //      
            RunZiWeiChartCalculations(context);
            DetermineMajorStars(context);
            DetermineAllAuxiliaryAndMinorStars(context);
            DetermineDecadeLuckCycles(context);
            DetermineLifeCycleStars(context);
            DetermineStarBrightness(context);
            UpdateFinalResult(context);
            return context.Result;
        }
        // 檔案: AstrologyService.cs

        // ... (在所有既有方法之後，或在類別底部新增)

        /// <summary>
        /// 呼叫底層邏輯，計算八字星煞並更新 context.Result
        /// </summary>
        private void DetermineBaziShensha(AstrologyCalculationContext context)
        {
            // 假設 DetermineCorrectBaziPillars 運行後，Bazi 資訊已存於 context.Result.Bazi
            if (context.Result.Bazi == null)
            {
                // 如果 Bazi 資訊尚未準備好，則跳過
                context.Result = context.Result with { BaziShensha = new List<string>() };
                return;
            }

            var bazi = context.Result.Bazi;

            // 呼叫實際計算邏輯
            var shenshaList = CalculateBaziShenshaLogic(bazi);

            // 因為 AstrologyChartResult 是 record，使用 'with' 關鍵字進行非破壞性更新
            context.Result = context.Result with
            {
                BaziShensha = shenshaList
            };
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
            if (yearZhi == 12 || yearZhi == 4 || yearZhi == 8) return 6; // 亥卯未 (12, 4, 8) -> 巳 (6)
            return 6; // 亥卯未 (12, 4, 8) -> 巳 (6)
        }
        // ==========================================================
        // 【修正方法】: 根據農曆月份計算天巫位置 (輸入為農曆月份 1-12)
        // 口訣：正五九月在巳(6)，二六十月在申(9)，三七十一在寅(3)，四八十二在亥(12)。
        // ==========================================================
        private int PlaceTienWu(int month)
        {
            // 農曆月份 1, 5, 9 -> 巳宮 (6)
            if (month == 1 || month == 5 || month == 9) return 6;

            // 農曆月份 2, 6, 10 -> 申宮 (9)
            if (month == 2 || month == 6 || month == 10) return 9;

            // 農曆月份 3, 7, 11 -> 寅宮 (3)
            if (month == 3 || month == 7 || month == 11) return 3;

            // 農曆月份 4, 8, 12 -> 亥宮 (12)
            if (month == 4 || month == 8 || month == 12) return 12;

            return 0;
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
        // 【新增方法】: 根據生年地支計算火星起點宮位
        // 口訣：申子辰人寅戌揚（火寅起，鈴戌起），寅午戌人丑卯方（火丑起，鈴卯起）...
        // 火星：申子辰 -> 寅(3) | 寅午戌 -> 丑(2) | 巳酉丑 -> 卯(4) | 亥卯未 -> 戌(11)
        // ==========================================================
        private int PlaceHuoXingStart(int yearZhi)
        {
            if (yearZhi == 9 || yearZhi == 1 || yearZhi == 5) return 3; // 申子辰 -> 寅(3)
            if (yearZhi == 3 || yearZhi == 7 || yearZhi == 11) return 2; // 寅午戌 -> 丑(2)
            if (yearZhi == 6 || yearZhi == 10 || yearZhi == 2) return 4; // 巳酉丑 -> 卯(4)
            if (yearZhi == 12 || yearZhi == 4 || yearZhi == 8) return 10; // 亥卯未 -> 戌(11)
            return 10; // 亥卯未 -> 戌(11)
        }

        // ==========================================================
        // 【新增方法】: 根據生年地支計算鈴星起點宮位
        // 口訣：申子辰人寅戌揚（火寅起，鈴戌起），寅午戌人丑卯方（火丑起，鈴卯起）...
        // 鈴星：申子辰 -> 戌(11) | 寅午戌 -> 卯(4) | 巳酉丑 -> 戌(11) | 亥卯未 -> 卯(4)
        // ==========================================================
        private int PlaceLingXingStart(int yearZhi)
        {
            if (yearZhi == 9 || yearZhi == 1 || yearZhi == 5) return 11; // 申子辰 -> 戌(11)
            if (yearZhi == 3 || yearZhi == 7 || yearZhi == 11) return 4; // 寅午戌 -> 卯(4)
            if (yearZhi == 6 || yearZhi == 10 || yearZhi == 2) return 11; // 巳酉丑 -> 戌(11)
            if (yearZhi == 12 || yearZhi == 4 || yearZhi == 8) return 11; // 亥卯未 -> 戌(11)
            return 4; // 亥卯未 -> 卯(4)
        }
        // ==========================================================
        // 【新增方法】: 根據生年天干計算天魁天鉞位置
        // 口訣：甲戊庚牛羊，乙己鼠猴鄉，丙丁豬雞位，壬癸兔蛇藏，辛年逢馬虎
        // ==========================================================
        private (int KuiPos, int YuePos) PlaceKuiYue(int yearGan)
        {
            switch (yearGan)
            {
                case 1: // 甲
                case 5: // 戊
                case 7: // 庚
                    return (2, 8); // 丑(牛), 未(羊)
                case 2: // 乙
                case 6: // 己
                    return (1, 9); // 子(鼠), 申(猴)
                case 3: // 丙
                case 4: // 丁
                    return (12, 10); // 亥(豬), 酉(雞)
                case 9: // 壬
                case 10: // 癸
                    return (4, 6); // 卯(兔), 巳(蛇)
                case 8: // 辛
                    return (7, 3); // 午(馬), 寅(虎)
                default:
                    return (0, 0);
            }
        }
        // ==========================================================
        // 【新增方法】: 根據生年干支計算旬空位置 (旬中空亡)
        // 口訣：從年干支順數至癸止，後兩位即是空亡。
        // ==========================================================
        private (int Pos1, int Pos2) PlaceXunKong(int yearGan, int yearZhi)
        {
            // Start at yearZhi (1-12) with yearGan (1-10).
            int currentZhi = yearZhi;
            int currentGan = yearGan;

            // Advance until the stem reaches 癸 (10)
            while (currentGan <= 10)
            {
                currentZhi = PalaceWrap(currentZhi + 1); // Move to the next branch
                currentGan++; // Move to the next stem
            }

            // currentZhi is now the branch immediately following 癸. This is the first empty branch.
            int pos1 = currentZhi;

            // The second empty branch is the one after the first.
            int pos2 = PalaceWrap(pos1 + 1);

            return (pos1, pos2);
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
            //if (yearGan % 2 != 0) { context.SecondaryStars[PalaceWrap(2 + (yearGan + 1) / 2 - 1)] += "魁 "; context.SecondaryStars[PalaceWrap(8 - ((yearGan + 1) / 2 - 1))] += "鉞 "; } else { context.SecondaryStars[PalaceWrap(9 - (yearGan / 2 - 1))] += "魁 "; context.SecondaryStars[PalaceWrap(3 + (yearGan / 2 - 1))] += "鉞 "; }
            // ------------------------------------------------------------------
            // 【修正】天魁、天鉞 (論生年干)
            // ------------------------------------------------------------------
            // ------------------------------------------------------------------
            // 【修正】天魁、天鉞 (論生年干)
            // ------------------------------------------------------------------
            var (tiKuiPos, tiYuePos) = PlaceKuiYue(yearGan);

            // 天魁 安入 SmallStars
            context.SecondaryStars[tiKuiPos] = (string.IsNullOrEmpty(context.SecondaryStars[tiKuiPos]) ? "" : context.SecondaryStars[tiKuiPos] + " ") + "魁";
            // 天鉞 安入 SmallStars
            context.SecondaryStars[tiYuePos] = (string.IsNullOrEmpty(context.SecondaryStars[tiYuePos]) ? "" : context.SecondaryStars[tiYuePos] + " ") + "鉞"; 
            int luCunPos = 0; switch (yearGan) { case 1: luCunPos = 3; break; case 2: luCunPos = 4; break; case 3: luCunPos = 6; break; case 4: luCunPos = 7; break; case 5: luCunPos = 6; break; case 6: luCunPos = 7; break; case 7: luCunPos = 9; break; case 8: luCunPos = 10; break; case 9: luCunPos = 12; break; case 10: luCunPos = 1; break; }
            context.LuCunPos = luCunPos;
            context.SecondaryStars[luCunPos] += "祿 "; context.BadStars[PalaceWrap(luCunPos + 1)] += "羊 "; context.BadStars[PalaceWrap(luCunPos - 1)] += "陀 ";
            int huoPosBase = new int[] { 2, 3, 4, 10, 2, 3, 4, 10, 2, 3, 4, 10 }[yearZhi - 1];
            int lingPosBase = new int[] { 4, 11, 11, 11, 4, 11, 11, 11, 4, 11, 11, 11 }[yearZhi - 1];
            // ------------------------------------------------------------------
            // 【修正】火星、鈴星 (論生年支、生時)
            // ------------------------------------------------------------------
            int huoStartPos = PlaceHuoXingStart(yearZhi);
            int lingStartPos = PlaceLingXingStart(yearZhi);

            // 火星：從起點順數至生時
            //int huoPos = PalaceWrap(huoStartPos + hourZhi - 1);
            // 鈴星：從起點順數至生時
            //int lingPos = PalaceWrap(lingStartPos + hourZhi - 1);

            //context.BadStars[huoPos] += "火 ";
            // 鈴星 安入 SmallStars
            //context.SmallStars[lingPos] = (string.IsNullOrEmpty(context.SmallStars[lingPos]) ? "" : context.SmallStars[lingPos] + " ") + "鈴";
            context.BadStars[PalaceWrap(huoStartPos + hour - 1)] += "火 "; context.BadStars[PalaceWrap(lingStartPos + hour - 1)] += "鈴 ";
            // context.BadStars[hourZhi] += "劫 "; context.BadStars[PalaceWrap(2 + 12 - hourZhi)] += "空 ";
            System.Action<string, string> AddFourTransformation = (starName, transType) => { for (int i = 1; i <= 12; i++) { string allStarsInPalace = context.CCM[i] + context.CCN[i] + context.SecondaryStars[i]; if (allStarsInPalace.Contains(starName)) { context.FourTransformationStars[i] += transType; return; } } };
            switch (yearGan) { case 1: AddFourTransformation("廉", "祿"); AddFourTransformation("破", "權"); AddFourTransformation("武", "科"); AddFourTransformation("陽", "忌"); break; case 2: AddFourTransformation("機", "祿"); AddFourTransformation("梁", "權"); AddFourTransformation("紫", "科"); AddFourTransformation("陰", "忌"); break; case 3: AddFourTransformation("同", "祿"); AddFourTransformation("機", "權"); AddFourTransformation("昌", "科"); AddFourTransformation("廉", "忌"); break; case 4: AddFourTransformation("陰", "祿"); AddFourTransformation("同", "權"); AddFourTransformation("機", "科"); AddFourTransformation("巨", "忌"); break; case 5: AddFourTransformation("貪", "祿"); AddFourTransformation("陰", "權"); AddFourTransformation("右", "科"); AddFourTransformation("機", "忌"); break; case 6: AddFourTransformation("武", "祿"); AddFourTransformation("貪", "權"); AddFourTransformation("梁", "科"); AddFourTransformation("曲", "忌"); break; case 7: AddFourTransformation("陽", "祿"); AddFourTransformation("武", "權"); AddFourTransformation("陰", "科"); AddFourTransformation("同", "忌"); break; case 8: AddFourTransformation("巨", "祿"); AddFourTransformation("陽", "權"); AddFourTransformation("曲", "科"); AddFourTransformation("昌", "忌"); break; case 9: AddFourTransformation("梁", "祿"); AddFourTransformation("紫", "權"); AddFourTransformation("左", "科"); AddFourTransformation("武", "忌"); break; case 10: AddFourTransformation("破", "祿"); AddFourTransformation("巨", "權"); AddFourTransformation("陰", "科"); AddFourTransformation("貪", "忌"); break; }

            string[] doctorStars = { "博", "力", "青", "小", "將", "奏", "飛", "吉", "病", "大", "伏", "官" };
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
            context.BadStars[PalaceWrap(7 - yearZhi + 1)] += "哭 "; context.BadStars[PalaceWrap(7 + (yearZhi - 1))] += "虛 ";
            context.GoodStars[PalaceWrap(5 + yearZhi - 1)] += "龍 "; context.GoodStars[PalaceWrap(11 - (yearZhi - 1))] += "鳳 ";
            int[] guPosMap = { 3, 3, 3, 6, 6, 6, 9, 9, 9, 12, 12, 12 }; int[] guaPosMap = { 11, 11, 11, 2, 2, 2, 5, 5, 5, 8, 8, 8 };
            context.BadStars[guPosMap[yearZhi - 1]] += "孤 "; context.BadStars[guaPosMap[yearZhi - 1]] += "寡 ";
            context.GoodStars[PalaceWrap(context.ZuoFuPos + day - 1)] += "三 "; context.GoodStars[PalaceWrap(context.YouBiPos - day + 1)] += "八 ";
            //少判斷文昌文曲在命宮或身宮時，不加恩貴
            context.GoodStars[PalaceWrap(WenChangPos + day - 2)] += "恩 ";  context.GoodStars[PalaceWrap(WenQuPos + day - 2)] += "貴 ";
            //if (WenQuPos != -1)
            //{
            //    // 1. 從文曲宮(wenQuPos)起算初一
            //    // 2. 順數到生日 (lunarDay)
            //    // 3. 退一步 (逆數一宮)
            //    int stepForward = day - 1; // 從初一(0步)到生日(lunarDay-1步)
            //    int tianGuiPos = (WenQuPos + stepForward - 1) % 12 + 1; // 順數到生日所在宮位

            //    // 再退一步 (逆時針一宮)
            //    tianGuiPos = tianGuiPos + 1;
            //    if (tianGuiPos == 0) tianGuiPos = 12;
            //    context.GoodStars[tianGuiPos] += (string.IsNullOrEmpty(context.GoodStars[tianGuiPos]) ? "" : "貴") ;
            //}
            // context.GoodStars[PalaceWrap(context.WenChangPos + day - 2)] += "恩 "; context.GoodStars[PalaceWrap(context.WenQuPos + day - 2)] += "貴 ";
            context.GoodStars[PalaceWrap(context.MingGongIndex + yearZhi - 1)] += "才 ";
            context.GoodStars[PalaceWrap(context.ShenGongIndex + yearZhi - 1)] += "壽 ";

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

            context.GoodStars[taiFuPos] += (string.IsNullOrEmpty(context.GoodStars[taiFuPos]) ? "" : " ") + "輔";
            context.GoodStars[fengGaoPos] += (string.IsNullOrEmpty(context.GoodStars[fengGaoPos]) ? "" : " ") + "誥";

            // ==========================================================
            // 【新增星曜的安星邏輯】
            // ==========================================================

            // 1. 天馬 (論年支)
            int tianMaPos = PlaceTienMa(yearZhi);
            context.GoodStars[tianMaPos] = (string.IsNullOrEmpty(context.GoodStars[tianMaPos]) ? "" : context.GoodStars[tianMaPos] + " ") + "馬";

            // 2. 天月 (論生月)
            int tianYuePos = PlaceTienYue(month);
            context.BadStars[tianYuePos] = (string.IsNullOrEmpty(context.BadStars[tianYuePos]) ? "" : context.BadStars[tianYuePos] + " ") + "月";

            // **17. 天巫 (Tien Wu) -> 巫** (修正：使用農曆月份 month)
            int tienWuPos = PlaceTienWu(month);
            context.BadStars[tienWuPos] = (string.IsNullOrEmpty(context.BadStars[tienWuPos]) ? "" : context.BadStars[tienWuPos] + " ") + "巫";

            // 3. 天哭、天虛 (論年支) (保留原邏輯並加入SmallStars)
            // 原有邏輯：
            // context.BadStars[PalaceWrap(7 + yearZhi - 1)] += "哭 "; context.BadStars[PalaceWrap(7 - (yearZhi - 1))] += "虛 ";
            // 修正為：
            int tianKuPos = PalaceWrap(7 + yearZhi - 1); // 午(7)起子年，順行年支 (天哭逆行) -> 應為 7 - (yearZhi - 1)
            int tianXuPos = PalaceWrap(7 - (yearZhi - 1)); // 午(7)起子年，逆行年支 (天虛順轉) -> 應為 7 + (yearZhi - 1)

            //// 根據「哭逆行兮虛順轉，數到生年便停留」的常見理解：
            //// 午宮起子年（子=1），哭逆轉，虛順轉
            //tianKuPos = PalaceWrap(7 + (yearZhi - 1));
            //tianXuPos = PalaceWrap(7 - (yearZhi - 1));

            //context.BadStars[tianKuPos] = (string.IsNullOrEmpty(context.BadStars[tianKuPos]) ? "" : context.BadStars[tianKuPos] + " ") + "哭";
            //context.BadStars[tianXuPos] = (string.IsNullOrEmpty(context.BadStars[tianXuPos]) ? "" : context.BadStars[tianXuPos] + " ") + "虛";

            // 4. (以截路空亡處理，論年干)
            var (jieLu1, jieLu2) = PlaceJieLuKongWang(yearGan);
            if (yearGan == 1 | yearGan == 3 | yearGan == 5 | yearGan == 7 | yearGan == 9 | yearGan == 11)
            {
                if (jieLu1 != 0) context.BadStars[jieLu1] = (string.IsNullOrEmpty(context.BadStars[jieLu1]) ? "" : context.BadStars[jieLu1] + " ") + "截";
            }
            else
            {
                if (jieLu2 != 0) context.BadStars[jieLu2] = (string.IsNullOrEmpty(context.BadStars[jieLu2]) ? "" : context.BadStars[jieLu2] + " ") + "截";
            }

            // ==========================================================
            // 【13. 旬空 (Xun Kong) -> 旬 (修正邏輯，放入 BadStars)】
            // ==========================================================
            var (xunKongPos1, xunKongPos2) = PlaceXunKong(yearGan, yearZhi);
            bool isYearGanYang = yearGan % 2 != 0;
            int xunKongTargetPos = 0;

            // 陽年空陽支 (陽支: 1, 3, 5, 7, 9, 11 - 奇數宮位)
            // 陰年空陰支 (陰支: 2, 4, 6, 8, 10, 12 - 偶數宮位)
            if (isYearGanYang)
            {
                // 陽年取陽支 (奇數)
                xunKongTargetPos = (xunKongPos1 % 2 != 0) ? xunKongPos1 : xunKongPos2;
            }
            else
            {
                // 陰年取陰支 (偶數)
                xunKongTargetPos = (xunKongPos1 % 2 == 0) ? xunKongPos1 : xunKongPos2;
            }
            if (xunKongTargetPos != 0)
            {
                context.BadStars[xunKongTargetPos] += "旬 ";
            }
            // 5. 破碎 (論年支)
            int poSuiPos = PlacePoSui(yearZhi);
            context.BadStars[poSuiPos] = (string.IsNullOrEmpty(context.BadStars[poSuiPos]) ? "" : context.BadStars[poSuiPos] + " ") + "碎";

            //// 6. 天壽 (論身宮，年支) (保留原邏輯並加入SmallStars)
            //// 原有邏輯： context.GoodStars[PalaceWrap(context.ShenGongIndex + day - 1)] += "壽 ";
            //// 修正為：
            //int tianShouPos = PalaceWrap(context.ShenGongIndex + yearZhi - 1); // 身宮起子年，順行年支
            //context.SmallStars[tianShouPos] = (string.IsNullOrEmpty(context.SmallStars[tianShouPos]) ? "" : context.SmallStars[tianShouPos] + " ") + "壽";

            //// 7. 天才 (論命宮，年支) (保留原邏輯並加入SmallStars)
            //// 原有邏輯： context.GoodStars[PalaceWrap(context.MingGongIndex + day - 1)] += "才 ";
            //// 修正為：
            //int tianCaiPos = PalaceWrap(context.MingGongIndex + yearZhi - 1); // 命宮起子年，順行年支
            //context.SmallStars[tianCaiPos] = (string.IsNullOrEmpty(context.SmallStars[tianCaiPos]) ? "" : context.SmallStars[tianCaiPos] + " ") + "才";

            // 8. 陰煞 (論生月)
            int yinShaPos = PlaceYinSha(month);
            context.BadStars[yinShaPos] = (string.IsNullOrEmpty(context.BadStars[yinShaPos]) ? "" : context.BadStars[yinShaPos] + " ") + "煞";

            // 9. 解神 (論生月)
            int jieShenPos = PlaceJieShen(month);
            context.GoodStars[jieShenPos] = (string.IsNullOrEmpty(context.GoodStars[jieShenPos]) ? "" : context.GoodStars[jieShenPos] + " ") + "解";
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

        public List<AnnualLuck> GenerateAnnualLucks(int fromYear, int toYear, string dayStem)
        {
            string[] GAN = { "甲", "乙", "丙", "丁", "戊", "己", "庚", "辛", "壬", "癸" };
            string[] ZHI = { "子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥" };
            var results = new List<AnnualLuck>();
            for (int y = fromYear; y <= toYear; y++)
            {
                int g = (y - 4) % 10;
                int z = (y - 4) % 12;
                var currStem = GAN[g];
                var currBranch = ZHI[z];
                results.Add(new AnnualLuck
                {
                    Year = y,
                    HeavenlyStem = currStem,
                    EarthlyBranch = currBranch,
                    StemLiuShen = GetLiuShen(dayStem, currStem),
                    BranchLiuShen = GetBranchLiuShenList(dayStem, currBranch),
                });
            }
            return results;
        }

        // 在 AstrologyService.cs 內新增 private 方法
        private List<string> GetAnnualInteractions(string annualBranch, List<string> baziBranches)
        {
            var interactions = new List<string>();
            var pillarNames = new[] { "年支", "月支", "日支", "時支" };

            for (int i = 0; i < 4; i++)
            {
                string baziBranch = baziBranches[i];
                string name = pillarNames[i];

                // --- 1. 六沖 (Six Clashes) ---
                if (GanZhiConstants.SixClashes.TryGetValue(annualBranch, out string clashTarget) && clashTarget == baziBranch ||
                    GanZhiConstants.SixClashes.TryGetValue(baziBranch, out string clashSource) && clashSource == annualBranch)
                {
                    interactions.Add($"與{name}相沖 ({annualBranch}沖{baziBranch})");
                }

                // --- 2. 六合 (Six Combinations) ---
                if (GanZhiConstants.SixCombinations.TryGetValue(annualBranch, out string combineTarget) && combineTarget == baziBranch ||
                    GanZhiConstants.SixCombinations.TryGetValue(baziBranch, out string combineSource) && combineSource == annualBranch)
                {
                    interactions.Add($"與{name}六合 ({annualBranch}合{baziBranch})");
                }

                // --- 3. 自刑 (Self-Punishment): 辰、午、酉、亥 見同支 ---
                if (annualBranch == baziBranch && ("辰午酉亥").Contains(annualBranch))
                {
                    interactions.Add($"與{name}自刑 ({annualBranch}刑{baziBranch})");
                }

                // ... (在此處添加三刑、三合、六害等邏輯)
            }

            return interactions.Distinct().ToList();
        }

        public List<string> GetBranchLiuShenList(string dayStem, string branch)
        {
            if (!BranchToHiddenStems.ContainsKey(branch)) return new List<string>();
            var result = new List<string>();
            foreach (var stem in BranchToHiddenStems[branch])
                result.Add(GetLiuShen(dayStem, stem));
            return result;
        }
        public void ExportAnnualLucksJson(List<AnnualLuck> list, string filename)
        {
            var json = JsonSerializer.Serialize(list, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(filename, json, Encoding.UTF8);
        }
        // 在 AstrologyService.cs 內新增 private 方法
        private List<string> CalculateBaziShensha(BaziInfo bazi)
        {
            var shenshaList = new List<string>();
            string dayStem = bazi.DayPillar.HeavenlyStem;
            string yearBranch = bazi.YearPillar.EarthlyBranch;

            var allBranches = new List<string>
            {
                bazi.YearPillar.EarthlyBranch,
                bazi.MonthPillar.EarthlyBranch,
                bazi.DayPillar.EarthlyBranch,
                bazi.TimePillar.EarthlyBranch
            };

            // --- 1. 天乙貴人 (以日干查) ---
            var tianYiTargets = new List<string>();
            if (new[] { "甲", "戊", "庚" }.Contains(dayStem)) tianYiTargets.AddRange(GanZhiConstants.ShenshaFromDayStem["天乙貴人_甲戊庚"]);
            else if (new[] { "乙", "己" }.Contains(dayStem)) tianYiTargets.AddRange(GanZhiConstants.ShenshaFromDayStem["天乙貴人_乙己"]);
            else if (new[] { "丙", "丁" }.Contains(dayStem)) tianYiTargets.AddRange(GanZhiConstants.ShenshaFromDayStem["天乙貴人_丙丁"]);
            else if (new[] { "壬", "癸" }.Contains(dayStem)) tianYiTargets.AddRange(GanZhiConstants.ShenshaFromDayStem["天乙貴人_壬癸"]);
            else if (dayStem == "辛") tianYiTargets.AddRange(GanZhiConstants.ShenshaFromDayStem["天乙貴人_辛"]);

            foreach (var branch in allBranches.Where(b => tianYiTargets.Contains(b)))
            {
                shenshaList.Add($"天乙貴人 ({branch}支)");
            }

            // --- 2. 華蓋 (以年支查) ---
            // 規則：寅午戌 (戌), 巳酉丑 (丑), 申子辰 (辰), 亥卯未 (未)
            string huaGaiTarget = "";
            if (new[] { "寅", "午", "戌" }.Contains(yearBranch)) huaGaiTarget = "戌";
            else if (new[] { "巳", "酉", "丑" }.Contains(yearBranch)) huaGaiTarget = "丑";
            else if (new[] { "申", "子", "辰" }.Contains(yearBranch)) huaGaiTarget = "辰";
            else if (new[] { "亥", "卯", "未" }.Contains(yearBranch)) huaGaiTarget = "未";

            foreach (var branch in allBranches.Where(b => b == huaGaiTarget))
            {
                shenshaList.Add($"華蓋 ({branch}支)");
            }

            // ... 依此類推，加入所有其他神煞的計算邏輯

            return shenshaList.Distinct().ToList();
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
            // 檢查 context.Result 是否包含 LunarBirthDate (它應該在 DetermineCorrectBaziPillars 步驟中獲得)
            //string lunarBirthDate = context.LunarBirthDate ?? "";
            var lunarDateString = $"{context.Calendar.ChineseYear}{context.Calendar.ChineseMonth}{context.LunarDay}";

            // 檢查 context.Result 是否包含 BaziShensha (它應該在 DetermineBaziShensha 步驟中計算好)
            List<string> baziShensha = context.Result.BaziShensha ?? new List<string>();
            context.Result = context.Result with
            {
                palaces = finalPalaces,
                WuXingJuText = context.WuXingJuText,
                MingZhu = context.MingZhu,
                ShenZhu = context.ShenZhu,
                // 【修正點】現在可以從 context 中取得值
                LunarBirthDate = lunarDateString, // <--- 將組合好的字串賦值
                BaziShensha = baziShensha,
            };
        }
        /// <summary>
        /// 實際的星煞計算邏輯，使用 BaziData.cs 中的 GanZhiConstants
        /// </summary>
        /// <summary>
        /// 實際的星煞計算邏輯，使用 BaziData.cs 中的 GanZhiConstants
        /// 注意：所有對常數的引用皆已修正為 ShenshaFromDayStem 或 ShenshaFromBranch
        /// </summary>
        /// <summary>
        /// 實際的星煞計算邏輯，使用 BaziData.cs 中的 GanZhiConstants
        /// </summary>
        // 檔案: AstrologyService.cs

        private List<string> CalculateBaziShenshaLogic(BaziInfo bazi)
        {
            var shenshaList = new List<string>();

            // =========================================================================
            // 1. 基礎資訊及輔助列表定義 (修正 'allStemsAndBranches' 作用域問題)
            // =========================================================================
            string dayStem = bazi.DayPillar.HeavenlyStem;
            string dayPillar = bazi.DayPillar.HeavenlyStem + bazi.DayPillar.EarthlyBranch;

            // 1.1 四柱地支 (用於日干查神煞)
            List<string> allBranches = new()
            {
                bazi.YearPillar.EarthlyBranch,
                bazi.MonthPillar.EarthlyBranch,
                bazi.DayPillar.EarthlyBranch,
                bazi.TimePillar.EarthlyBranch
            };

            // 1.2 四柱干支總覽 (用於月支查神煞，因月神煞目標可能為天干)
            List<string> allStemsAndBranches = new()
            {
                bazi.YearPillar.HeavenlyStem, bazi.YearPillar.EarthlyBranch,
                bazi.MonthPillar.HeavenlyStem, bazi.MonthPillar.EarthlyBranch,
                bazi.DayPillar.HeavenlyStem, bazi.DayPillar.EarthlyBranch,
                bazi.TimePillar.HeavenlyStem, bazi.TimePillar.EarthlyBranch
            };


            // =========================================================================
            // 【2. 從日干 (Day Stem) 查神煞】(使用 allBranches)
            // =========================================================================

            // --- 通用查找方法 (適用於 天乙、文昌、天廚、學堂、紅艶、金輿、祿神、學士、福星、國印、太極、流霞) ---
            var dayShenshaMappings = new (string Name, string KeyPrefix, bool IsGroupKey)[]
            {
                ("天乙貴人", dayStem, false),
                ("文昌貴人", "文昌", false),
                ("天廚", "天廚", false),
                ("學堂", "學堂", false),
                ("紅艶", "紅艶", false),
                ("金輿", "金舆", false),
                ("祿神", "祿神", false),
                ("學士", "學士", false),
                ("福星", "福星", false),
                ("國印", "國印", false),
                ("太極貴人", "太極", true),
                ("流霞", "流霞", false),
            };

            foreach (var mapping in dayShenshaMappings)
            {
                string lookupKey = "";

                if (mapping.Name == "天乙貴人")
                {
                    lookupKey = dayStem; // 天乙貴人直接用日干
                }
                else if (mapping.IsGroupKey) // 太極貴人 (Group Key Logic)
                {
                    lookupKey = dayStem switch
                    {
                        "甲" or "乙" => $"{mapping.KeyPrefix}_甲",
                        "丙" or "丁" => $"{mapping.KeyPrefix}_丙",
                        "戊" or "己" => dayStem == "戊" ? $"{mapping.KeyPrefix}_戊" : $"{mapping.KeyPrefix}_己", // 特殊處理太極貴人戊己
                        "庚" or "辛" => $"{mapping.KeyPrefix}_庚",
                        "壬" or "癸" => $"{mapping.KeyPrefix}_壬",
                        _ => ""
                    };
                }
                else // 其他單干神煞 (如 文昌_甲, 天廚_乙)
                {
                    lookupKey = $"{mapping.KeyPrefix}_{dayStem}";
                }

                // 修正文昌/天廚的特殊合併規則，確保查詢正確
                if (mapping.Name == "文昌貴人" || mapping.Name == "天廚")
                {
                    lookupKey = dayStem switch
                    {
                        "戊" => $"{mapping.KeyPrefix}_丙", // 戊對應丙的文昌/天廚地支
                        "己" => $"{mapping.KeyPrefix}_丁", // 己對應丁的文昌/天廚地支
                        _ => $"{mapping.KeyPrefix}_{dayStem}"
                    };
                }

                // 修正 lookupKey 的最終值
                if (mapping.Name != "天乙貴人" && !mapping.IsGroupKey)
                {
                    lookupKey = $"{mapping.KeyPrefix}_{dayStem}";
                }

                if (GanZhiConstants.ShenshaFromDayStem.TryGetValue(lookupKey, out string[]? targets))
                {
                    foreach (var branch in allBranches)
                    {
                        // 排除流霞中非地支的項目 (如: 乙干對應的'戊')
                        if (targets.Contains(branch) && GanZhiConstants.Zhi.Contains(branch))
                        {
                            // 找到該地支出現在哪一柱
                            string pillarName = GetPillarName(branch, bazi);

                            // 將結果格式化為: "金輿 (亥支 - 日柱)"
                            shenshaList.Add($"{mapping.Name} ({branch}支 - {pillarName}柱)");
                            //shenshaList.Add($"{mapping.Name} ({branch}支)");
                        }
                    }
                }
            }


            // --- 羊刃 (Yang Ren) (特殊處理：多個地支) ---
            string yangRenKey = $"羊刃_{dayStem}";
            if (GanZhiConstants.ShenshaFromDayStem.TryGetValue(yangRenKey, out string[]? yangRenTargets))
            {
                foreach (var branch in allBranches.Where(yangRenTargets.Contains))
                {
                    // 找到該地支出現在哪一柱
                    string pillarName = GetPillarName(branch, bazi);
                    shenshaList.Add($"羊刃 ({branch}支- {pillarName}柱");
                }
            }

            // --- 掃妻 (Sao Qi) (特殊：只看日干，對應的結果為日柱地支) ---
            string saoQiKey = $"掃妻_{dayStem}";
            if (GanZhiConstants.ShenshaFromDayStem.TryGetValue(saoQiKey, out string[]? saoQiTargets))
            {
                string saoQiTargetBranch = saoQiTargets.First(); // 只有一個目標地支
                if (bazi.DayPillar.EarthlyBranch == saoQiTargetBranch)
                {
                    shenshaList.Add($"掃妻 ({dayPillar}日柱)");
                }
            }

            // --- 魁罡 (Kui Gang) (特殊：只看日柱干支組合) ---
            if (new[] { "庚辰", "庚戌", "壬辰", "戊戌" }.Contains(dayPillar))
            {
                shenshaList.Add("魁罡 (日柱)");
            }
            // --- 金神 (特殊：只看日柱干支組合) ---
            if (new[] { "乙丑", "己巳", "癸酉" }.Contains(dayPillar))
            {
                shenshaList.Add("金神 (日柱)");
            }

            // --- 十惡大敗   (特殊：只看日柱干支組合) ---
            if (new[] { "甲辰","乙巳","丙申","丁亥","戊戌","己丑","庚辰","辛巳","壬申","癸亥" }.Contains(dayPillar))
            {
                shenshaList.Add("十惡大敗 (日柱)");
            }
            // 陰差陽錯
            if (new[] { "丙子", "丁丑", "戊寅", "辛卯", "壬辰", "癸巳", "丙午", "丁未", "戊申", "辛酉", "壬戌", "癸亥" }.Contains(dayPillar))
            {
                shenshaList.Add("陰差陽錯 (日柱)");
            }
            // 八專
            if (new[] { "甲寅", "乙卯", "丁未", "戊戌", "己未", "庚申", "辛酉", "癸丑" }.Contains(dayPillar))
            {
                shenshaList.Add("八專 (日柱)");
            }
            // 九醜
            if (new[] { "戊子","戊午","壬子","壬午","丁卯","己卯","辛卯","丁酉","己酉","辛酉" }.Contains(dayPillar))
            {
                shenshaList.Add("九醜 (日柱)");
            }

            // =========================================================================
            // 【3. 從月支 (Month Branch) 查神煞】(使用 allStemsAndBranches)
            // =========================================================================

            string monthBranch = bazi.MonthPillar.EarthlyBranch;
            string yearBranch = bazi.YearPillar.EarthlyBranch; // 舊的年支查神煞可能還需要
            string dayBranch = bazi.DayPillar.EarthlyBranch;

            // --- 輔助函式：用於月支查神煞 ---
            var monthShenshaMappings = new string[]
            {
                "天德貴", "月德貴", "月空", "將軍", "德貴", "秀貴", "天德",
                "月德", "天醫", "陰陽", "斷橋", "牢獄"
            };

            foreach (var name in monthShenshaMappings)
            {
                string lookupKey = $"{name}_{monthBranch}";

                if (GanZhiConstants.ShenshaFromMonthBranch.TryGetValue(lookupKey, out string[]? targets))
                {
                    foreach (var element in allStemsAndBranches)
                    {
                        if (targets.Contains(element))
                        {
                            // 區分是天干還是地支
                            // 由於 targets 內可能有 '丙丁' 這種組合，我們只檢查 targets 內是否包含單個干或支
                            string location = GanZhiConstants.Gan.Contains(element) ? "干" : (GanZhiConstants.Zhi.Contains(element) ? "支" : "N/A");

                            if (location != "N/A")
                            {
                                // 找到該干支位於哪一柱 (可選，但讓輸出更詳細)
                                string pillarName = GetPillarName(element, bazi);

                                shenshaList.Add($"{name} ({element}{location} - {pillarName}柱)");
                            }
                        }
                    }
                }
            }
            // --- 輔助函式：用於日支查神煞 ---
            var dayBranchShenshaMappings = new string[]
            {
                "驛馬", "華蓋", "將星", "亡神", "劫煞",
                "孤辰", "寡宿", "咸池", "天喜", "紅鸞"
            };

            foreach (var name in dayBranchShenshaMappings)
            {
                string lookupKey = $"{name}_{dayBranch}";

                if (GanZhiConstants.ShenshaFromDayBranchLookup.TryGetValue(lookupKey, out string[]? targets))
                {
                    foreach (var branch in allBranches)
                    {
                        if (targets.Contains(branch))
                        {
                            // 找到該地支位於哪一柱
                            string pillarName = GetPillarName(branch, bazi);

                            // 咸池通常又稱桃花
                            string displayName = name == "咸池" ? "咸池(桃花)" : name;

                            // 如果神煞所在之地支與發動神煞的日支相同，則明確標註為「日支查自身」
                            string locationDetails = (branch == dayBranch)
                                ? $"{branch}支 - 日支查自身"
                                : $"{branch}支 - {pillarName}柱";

                            shenshaList.Add($"{displayName} (日支起查: {locationDetails})");
                        }
                    }
                }
            }
            // =========================================================================
            // 【2. 從年支 (Year Branch) 查神煞】 (保持不變)
            // ... (此處省略年支查神煞的邏輯，請保持您原有的實現或根據需求補齊)
            // =========================================================================
            // =========================================================================
            // 【4. 從年支 (Year Branch) 查神煞】(使用 allBranches)
            // =========================================================================

            // --- 輔助函式：用於年支查神煞 ---
            var yearShenshaMappings = new string[]
            {
                "驛馬", "華蓋", "將星", "亡神", "劫煞", "災煞",
                "孤辰", "寡宿", "咸池", "天喜", "紅鸞", "喪門", "白虎"
            };

            foreach (var name in yearShenshaMappings)
            {
                string lookupKey = $"{name}_{yearBranch}";

                if (GanZhiConstants.ShenshaFromYearBranch.TryGetValue(lookupKey, out string[]? targets))
                {
                    foreach (var branch in allBranches)
                    {
                        if (targets.Contains(branch))
                        {
                            // 找到該地支位於哪一柱
                            string pillarName = GetPillarName(branch, bazi);

                            // 咸池通常又稱桃花
                            string displayName = name == "咸池" ? "咸池(桃花)" : name;

                            shenshaList.Add($"{displayName} ({branch}支 - {pillarName}柱)");
                        }
                    }
                }
            }
            // =========================================================================
            // 【6. 計算空亡 (Kong Wang) 】(以年柱和日柱為準)
            // =========================================================================

            // 6.1 年柱空亡判斷
            string yearPillar = bazi.YearPillar.HeavenlyStem + bazi.YearPillar.EarthlyBranch;
            string yearXunShou = FindXunShou(yearPillar);
            if (yearXunShou != null && GanZhiConstants.KongWangRules.TryGetValue(yearXunShou, out string[]? yearKongWangBranches))
            {
                CheckAndAddKongWang("年柱空亡", yearKongWangBranches, bazi, shenshaList);
            }

            // 6.2 日柱空亡判斷
            //string dayPillar = bazi.DayPillar.HeavenlyStem + bazi.DayPillar.EarthlyBranch;
            string dayXunShou = FindXunShou(dayPillar);
            if (dayXunShou != null && GanZhiConstants.KongWangRules.TryGetValue(dayXunShou, out string[]? dayKongWangBranches))
            {
                CheckAndAddKongWang("日柱空亡", dayKongWangBranches, bazi, shenshaList);
            }
            // 確保結果不重複
            return shenshaList.Distinct().ToList();
        }
        private string FindXunShou(string ganZhi)
        {
            // 假設 GanZhiConstants.All60GanZhi 包含了完整的 60 甲子順序
            int index = Array.IndexOf(GanZhiConstants.All60GanZhi, ganZhi);

            if (index == -1) return null;

            // 每 10 個干支為一旬: 0-9 為甲子旬，10-19 為甲戌旬，以此類推
            int cycleIndex = index / 10;

            // 旬首固定為六甲旬的六個開頭
            return cycleIndex switch
            { 0 => "甲子", 1 => "甲戌", 2 => "甲申", 3 => "甲午",4 => "甲辰",5 => "甲寅",_ => null };
        }

        // Ecanapi.Services/AstrologyService.cs

        // ... (FindCsvRowAsync 保持不變)

        /// <summary>
        /// 非同步載入所有分析所需的 CSV 數據。
        /// </summary>
        private async Task<CsvDataContainer> LoadRequiredDataAsync(BaziInfo bazi)
        {
            var container = new CsvDataContainer();

            // 1. 定義四柱的干支組合和它們的名稱
            var pillars = new List<(string Name, string GanZhi)>
            {
                // 柱位名稱 (Key) | 干支組合 (Value)
                ("YearPillar", bazi.YearPillar.HeavenlyStem + bazi.YearPillar.EarthlyBranch),
                ("MonthPillar", bazi.MonthPillar.HeavenlyStem + bazi.MonthPillar.EarthlyBranch),
                ("DayPillar", bazi.DayPillar.HeavenlyStem + bazi.DayPillar.EarthlyBranch),
                ("TimePillar", bazi.TimePillar.HeavenlyStem + bazi.TimePillar.EarthlyBranch)
            };

            // 查詢的 CSV 檔案資訊
            // ⭐ 注意：您的範例是 "六十甲子.csv"，請確認正確的檔名是否為 "data-六十甲子.csv"
            string csvFileName = "六十甲子.csv";
            string searchColumn = "rgz"; // 查詢欄位是 rgz (干支)

            // --- 查詢 data-六十甲子.csv (四次查詢) ---
            foreach (var pillar in pillars)
            {
                string pillarName = pillar.Name;
                string ganZhi = pillar.GanZhi;

                // 執行非同步查詢
                var rowData = await FindCsvRowAsync(csvFileName, searchColumn, ganZhi);

                // 儲存結果到 container.LiuShiJiaZiData
                if (rowData != null)
                {
                    // 將結果以 "YearPillar" 或 "DayPillar" 作為 Key 存入字典
                    container.LiuShiJiaZiData.Add(pillarName, rowData);
                }
                // 如果找不到，該 Key 不會存在於字典中，分析器在使用時需檢查 Key 是否存在。
            }

            // --- 處理其他 CSV (如 data-干支組合.csv) ...
            // ...

            return container;
        }

        // Ecanapi.Services/AstrologyService.cs

        // ... (在 AstrologyService 類別內)

        /// <summary>
        /// 依據指定的檔案名稱、搜尋欄位和搜尋值，從 CSV 檔案中查找單一資料列。
        /// </summary>
        /// <param name="fileName">CSV 檔案名稱，例如 "data-六十甲子.csv" 或 "data-干支組合.csv"。</param>
        /// <param name="searchColumnName">要搜尋的欄位名稱 (標題列)。</param>
        /// <param name="searchValue">要搜尋的欄位值。</param>
        /// <returns>回傳符合條件的資料列，鍵為欄位名稱，值為資料，找不到則回傳 null。</returns>
        // ... 保持 private 即可

        // Ecanapi.Services/AstrologyService.cs (AstrologyService 類別內部)

        // 假設您的 IWebHostEnvironment 已經注入到 AstrologyService 中
        //private readonly IWebHostEnvironment _env;

        // ... (其他方法和屬性)

        private async Task<Dictionary<string, string>> FindCsvRowAsync(
           string fileName,
           string searchColumnName,
           string searchValue)
        {
            // 1. 構建路徑 (沿用 BaseDirectory 修正)
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string csvFilePath = Path.Combine(baseDirectory, "Data", fileName);

            if (!File.Exists(csvFilePath))
            {
                throw new FileNotFoundException($"CSV 檔案不存在於: {csvFilePath}");
            }

            // 2. 讀取所有行
            string[] lines = await File.ReadAllLinesAsync(csvFilePath);
            if (lines.Length <= 1) return null; // 至少需要標題行和一行資料

            // 3. 獲取標題行
            string headerLine = lines[0].Trim();
            var headers = headerLine.Trim('"')
                                    .Split(new[] { "\",\"" }, StringSplitOptions.None)
                                    .ToList();

            // 儲存欄位索引
            int searchColumnIndex = headers.IndexOf(searchColumnName);

            // 4. 檢查欄位是否存在
            if (searchColumnIndex == -1)
            {
                throw new InvalidOperationException($"CSV 檔案標題遺失查詢欄位: {searchColumnName}");
            }

            // 5. 尋找目標行並合併欄位 (修正跨行邏輯)
            for (int i = 1; i < lines.Length; i++)
            {
                string currentRecord = lines[i];

                // ⭐【核心修正 A】：跨行資料合併邏輯
                // 檢查：如果當前記錄的引號數是奇數，則表示是一個跨行欄位，需要繼續合併下一行。
                // 這個邏輯比檢查 StartWith/EndWith 更加通用。
                int quoteCount = currentRecord.Count(c => c == '"');

                while (quoteCount % 2 != 0 && i < lines.Length - 1)
                {
                    i++; // 移動到下一行
                    currentRecord += "\r\n" + lines[i]; // 合併下一行，並用換行符分隔（保留資料內部的換行）
                    quoteCount = currentRecord.Count(c => c == '"'); // 重新計算引號數量
                }

                // currentRecord 現在是一條完整的資料記錄。

                // ⭐【核心修正 B】：精準分割與清除邏輯

                // 1. 移除整行最外層引號（如果存在）
                string innerContent = currentRecord.Trim();
                if (innerContent.StartsWith('"') && innerContent.EndsWith('"'))
                {
                    // 注意：我們只從頭尾移除一個引號
                    innerContent = innerContent.Substring(1, innerContent.Length - 2);
                }

                // 2. 使用 '","' 分隔符號進行分割 (此時 innerContent 應為：1","丙戍","熱情豪爽...)
                var values = innerContent.Split(new[] { "\",\"" }, StringSplitOptions.None).ToList();

                // 6. 比對搜尋欄位的值
                if (values.Count > searchColumnIndex)
                {
                    string columnValue = values[searchColumnIndex]
                        .Trim()      // 1. 移除外部空白
                        .Trim('"')   // 2. 移除欄位兩側的引號（如果存在）
                        .Trim();     // 3. 再次移除引號內可能存在的空白字元

                    // 進行比對
                    if (columnValue == searchValue)
                    {
                        // 找到目標行，構建結果字典
                        var result = new Dictionary<string, string>();
                        for (int j = 0; j < headers.Count; j++)
                        {
                            // 檢查 values 數量是否匹配，避免索引超出範圍
                            if (j < values.Count)
                            {
                                // 對所有值進行相同的精準清除
                                string finalValue = values[j]
                                    .Trim()
                                    .Trim('"')
                                    .Trim();
                                result.Add(headers[j], finalValue);
                            }
                        }
                        return result;
                    }
                }
                // 如果 values 數量不足，或比對失敗，繼續到下一筆記錄 (外層 for 迴圈)
            }

            return null; // 找不到符合條件的行
        }

        // 輔助方法：檢查空亡地支並新增到列表
        private void CheckAndAddKongWang(string kwNamePrefix, string[] kwBranches, BaziInfo bazi, List<string> shenshaList)
        {
            // 包含所有四柱地支及其名稱
            var allPillars = new (string Branch, string PillarName)[]
            {
                (bazi.YearPillar.EarthlyBranch, "年"),
                (bazi.MonthPillar.EarthlyBranch, "月"),
                (bazi.DayPillar.EarthlyBranch, "日"),
                (bazi.TimePillar.EarthlyBranch, "時")
            };

            foreach (var (branch, pillarName) in allPillars)
            {
                if (kwBranches.Contains(branch))
                {
                    // 範例輸出: 年柱空亡 (戌支 - 月柱)
                    shenshaList.Add($"{kwNamePrefix} ({branch}支 - {pillarName}柱)");
                }
            }
        }
        // --- 輔助函式：用來判斷干支在四柱中的位置 (新增此方法到 AstrologyService 類別中) ---
        private string GetPillarName(string element, BaziInfo bazi)
        {
            // 檢查地支
            if (bazi.YearPillar.EarthlyBranch == element) return "年";
            if (bazi.MonthPillar.EarthlyBranch == element) return "月";
            if (bazi.DayPillar.EarthlyBranch == element) return "日";
            if (bazi.TimePillar.EarthlyBranch == element) return "時";

            // 檢查天干 (如果神煞是以天干為依據)
            if (bazi.YearPillar.HeavenlyStem == element) return "年";
            if (bazi.MonthPillar.HeavenlyStem == element) return "月";
            if (bazi.DayPillar.HeavenlyStem == element) return "日";
            if (bazi.TimePillar.HeavenlyStem == element) return "時";

            return ""; // 如果找不到，回傳空字串
        }  
    }
}