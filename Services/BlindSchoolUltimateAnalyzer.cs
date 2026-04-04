// Services/BlindSchoolUltimateAnalyzer.cs
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ecanapi.Models; // 【重要】引用模型所在的命名空間

namespace Ecanapi.Services
{
    /// <summary>
    /// 盲派終極三層合一分析引擎。
    /// 包含所有的基礎字典、核心判斷邏輯和分析方法。
    /// </summary>
    public class BlindSchoolUltimateAnalyzer
    {

        #region 基礎字典（請從您的「終極入口：三層合一分析.txt」複製所有字典）

        // 示例字典 1: 五行對應
        private static readonly Dictionary<char, string> WX = new()
        {
            {'甲',"木"},{'乙',"木"},{'丙',"火"},{'丁',"火"},{'戊',"土"},{'己',"土"},{'庚',"金"},{'辛',"金"},{'壬',"水"},{'癸',"水"}
        };
        private static readonly Dictionary<char, int> WN = new()
        {
            {'甲',1},{'乙',2},{'丙',3},{'丁',4},{'戊',5},{'己',6},{'庚',7},{'辛',8},{'壬',9},{'癸',10}
        };

        // 示例字典  四柱對應
        private static readonly Dictionary<int, string> zMst = new()
        {
            {0,"年"},{1,"月"},{2,"日"},{3,"時"}
        };

        // 示例字典 2: 地支藏干
        private static readonly Dictionary<char, List<(char Gan, string Level)>> CangGan = new()
        {
          {'子', new(){('癸',"本")}}, {'丑', new(){('己',"本"),('癸',"餘"),('辛',"餘")}},
         {'寅', new(){('甲',"本"),('丙',"中"),('戊',"餘")}}, {'卯', new(){('乙',"本")}},
        {'辰', new(){('戊',"本"),('乙',"餘"),('癸',"餘")}}, {'巳', new(){('丙',"本"),('庚',"中"),('戊',"餘")}},
        {'午', new(){('丁',"本"),('己',"餘")}}, {'未', new(){('己',"本"),('丁',"餘"),('乙',"餘")}},
        {'申', new(){('庚',"本"),('壬',"中"),('戊',"餘")}}, {'酉', new(){('辛',"本")}},
        {'戌', new(){('戊',"本"),('辛',"餘"),('丁',"餘")}}, {'亥', new(){('壬',"本"),('甲',"中")}}
        };
        private static readonly Dictionary<char, char> Sheng = new() { { '木', '火' }, { '火', '土' }, { '土', '金' }, { '金', '水' }, { '水', '木' } };
        private static readonly Dictionary<char, char> Ke = new() { { '木', '土' }, { '土', '水' }, { '水', '火' }, { '火', '金' }, { '金', '木' } };
        private static readonly int[] BranchPower = { 1, 4, 2, 3 }; // 年月日時
                                                                    // ... (請從您的文字檔中複製所有其他字典，如十神對應、格局判斷字典等) ...

        #endregion
        /// <summary>
        /// 主函數：判斷某十神是否為「真神」或「假神」
        /// </summary>
        /// <param name="year">年柱：如 "甲子"</param>
        /// <param name="month">月柱</param>
        /// <param name="day">日柱（日主）</param>
        /// <param name="hour">時柱</param>
        /// <param name="shiShen">要判斷的十神天干（如 '丙'）</param>
        /// <returns>判斷結果物件</returns>
        public ShiShenResult AnalyzeShiShen(string year, string month, string day, string hour, char shiShen, CsvDataContainer csvData)
        {
            // ⭐【取出資料】：從結果物件中存取 CSV 數據
            var retrievedCsvData = csvData;

            // 範例：取出日柱的六十甲子分析
            var dayPillarData = retrievedCsvData.LiuShiJiaZiData["DayPillar"];

            var pillars = new[] { year, month, day, hour };
            char riGan = day[0];
            string riWuXing = WX[riGan];

            var result = new ShiShenResult
            {
                ShiShen = shiShen.ToString(),
                WuXing = WX[shiShen],
                IsTrueGod = false,
                IsReal = false,
                RootType = "",
                SitBranchPower = 0,
                Explanation = new List<string>()
            };
            result.ShiShen = "";
            // Step 1: 日干判斷真假神（看地支根基）
            result = JudgeTrueFalseGod(pillars, riGan, result);

            // Step 2: 日干判斷虛實（看坐支力量）
            result = JudgeRealVirtual(pillars, riGan, result);

            // Step 3: 日干綜合現象解釋
            //result.Phenomenon = GetPhenomenon(shiShen, riGan, result.IsTrueGod, result.IsReal);

            return result;
        }

        // === Step 1: 真假神判斷 ===
        private static ShiShenResult JudgeTrueFalseGod(string[] pillars, char Sgan, ShiShenResult r)
        {
            char dayGan = pillars[2][0];
            char targetWX = pillars[0][0];
            bool hasRoot = false;
            string rootLevel = "";
            r.RootType = "";
            r.Gan = dayGan;
            int j = 0;
            for (int i = 0; i < 4; i++)
            {
                foreach (string pillar in pillars)
                {                    
                    //依年月日時順序地支
                    char zhi = pillar[1];
                    //取得藏干
                    var cang = CangGan[zhi];

                    foreach (var (gan, level) in cang)
                    {
                        if (pillars[i][0] == gan)
                        {
                            hasRoot = true;
                            if (level == "本") rootLevel = rootLevel + zMst[i] + pillars[i][0] + GetShiShenName(dayGan, pillars[i][0]) + "在" + zMst[j]+ zhi + "支" + "本氣根（最強）;";
                            else if (level == "中" ) rootLevel = rootLevel + zMst[i] + pillars[i][0] + GetShiShenName(dayGan, pillars[i][0]) + "在" + zMst[j] + zhi + "支" + "中氣根;";
                            else if (level == "餘" ) rootLevel = rootLevel + zMst[i] + pillars[i][0] + GetShiShenName(dayGan, pillars[i][0]) + "在" + zMst[j] + zhi + "支" + "餘氣根（較弱）;";
                        }
                        else
                        {
                            rootLevel = rootLevel + "";//  + zMst[i] + zhi + "支" + ""; //"無根;";
                        }
                    }
                    if (i == 0)
                    {
                        r.ShiShen = r.ShiShen + zMst[j]+ pillar[0] + GetShiShenName(dayGan, pillar[0]) + zhi+"-";
                    }
                    if (i==2) r.IsTrueGod = hasRoot;

                    r.Explanation.Add(hasRoot ? zMst[i] + pillars[i][0] + zMst[j] + zhi + "支" + "✔ 有根 → 真神" : zMst[i] + pillars[i][0] + zMst[j] + zhi + "支" + "✘ 無根 → 假神");
                    j++;
                }
                // 修正：foreach 結束後只 append 一次，避免重複累加
                if (hasRoot)
                    r.RootType = r.RootType + rootLevel;
                j = 0;
                hasRoot = false;
                rootLevel = "";
            }

            return r;
        }

        // === Step 2: 虛實判斷（坐支力量）===
        private static ShiShenResult JudgeRealVirtual(string[] pillars, char shiShen, ShiShenResult r)
        {
            //找出月支透干
            //依年月日時順序地支
            char zhi = pillars[1][1];
            //取得藏干
            var cang = CangGan[zhi];
            var positions = new List<string>();
            foreach (var (gan, level) in cang)
            {
                for (int i = 0; i < 4; i++)
                {
                    if (level == "本") r.Phenomenon = GetStyle(GetShiShenName(shiShen, gan));
                    if (pillars[i][0] == gan)
                    {
                        if (level == "本")
                        {
                            r.SitBranchPower = 1.5;
                            positions.Add(GetStyle(GetShiShenName(shiShen, gan)));
                            r.Phenomenon = GetStyle(GetShiShenName(shiShen, gan));
                        }
                        else if (level == "中")
                        {
                            r.SitBranchPower = r.SitBranchPower + 1.3;
                            positions.Add(GetStyle(GetShiShenName(shiShen, gan)));
                            r.Phenomenon = r.Phenomenon + GetStyle(GetShiShenName(shiShen, gan));
                        }
                        else if (level == "餘")
                        {
                            r.SitBranchPower = r.SitBranchPower + 1.1;
                            positions.Add(GetStyle(GetShiShenName(shiShen, gan)));
                            r.Phenomenon = r.Phenomenon + GetStyle(GetShiShenName(shiShen, gan));
                        }
                    }
                }

            }
            return r;
        }
            // 找出透出 shiShen 的所有天干位置
            //var positions = new List<int>();
            //for (int i = 0; i < 4; i++)
            //{
            //    if (pillars[i][0] == shiShen) positions.Add(i);
            //}

            //if (!positions.Any())
            //{
            //    r.IsReal = false;
            //    r.SitBranchPower = 0;
            //    r.Explanation.Add("✘ 月支未透干 → 無法判坐支 → 虛");
            //    return r;
            //}
            //else
            //return r;
            //}
                //int maxPower = 0;
                //foreach (int pos in positions)
                //{
                //    int power = BranchPower[pos]; // 月支4, 時支3, 年1, 日2
                //    if (power > maxPower) maxPower = power;
                //}

            //    double maxPower = 0;
            //        foreach (double pos in positions)
            //        {
            //            maxPower = maxPower + pos; // 月支4, 時支3, 年1, 日2
            //        }
            //        if (maxPower > 1)
            //        {
            //            //天干有根定格局
            //        }
            //        else
            //        {
            //            //天干無根以月支本氣定格局
            //        }

            //r.SitBranchPower = maxPower;
            //r.IsReal = maxPower >= 3; // 月支或時支透出 = 實
            //r.Explanation.Add(r.IsReal ? $"✔ 坐支力量 {maxPower} → 實" : $"✘ 坐支力量 {maxPower} → 虛");

            //return r;
        //}

        // === Step 3: 十神現象對照表 ===
        private static string GetPhenomenon(char shiShen, char riGan, bool isTrue, bool isReal)
        {
            string shiShenName = GetShiShenName(riGan, shiShen);

            if (!isTrue && !isReal) return $"【{shiShenName}】100%虛假 → 純粹表面、騙局、泡沫";

            return shiShenName switch
            {
                "正官" => isTrue ? "正官:事業真實穩定，有權威" : "官職虛名，易被貶",
                "七殺" => isTrue ? "七殺:權力真實，掌控力強" : "假威風，易惹官司",
                "正財" => isTrue ? "正財:財源穩定，正當收入" : "易被騙、傳銷、假投資",
                "偏財" => isTrue ? "偏財:橫財真實，機會好" : "賭博輸大錢",
                "正印" => isTrue ? "正印:學識真才，貴人實助" : "文憑假、合同騙",
                "偏印" => isTrue ? "偏印:偏門技術真" : "巫術騙局",
                "比肩" => isTrue ? "比肩:朋友講義氣，自力強" : "表面兄弟，實際坑你",
                "劫財" => isTrue ? "劫財:合伙真仗義" : "朋友借錢不還",
                "食神" => isTrue ? "食神:才藝真實，享受生活" : "嘴上功夫，無實績",
                "傷官" => isTrue ? "傷官:創新真才華" : "驕傲自大，無人用",
                _ => "未知"
            };
        }

        // === Step 3: 十神現象對照表 ===
        private static string GetStyle(string shiShen)
        {
            string shiShenName = "";
            switch (shiShen)
            {
                case "正官":  shiShenName = "正官格:事業真實穩定，有權威"; break;
                case "七殺":  shiShenName = "七殺格:權力真實，掌控力強"; break; 
                case "正財":  shiShenName = "正財格:財源穩定，正當收入"; break;
                case "偏財":  shiShenName = "偏財格:橫財真實，機會好"; break;
                case "正印":  shiShenName = "正印格:學識真才，貴人實助"; break;
                case "偏印":  shiShenName = "偏印格:偏門技術真";  break;
                case "比肩":  shiShenName = "比肩格:朋友講義氣，自力強"; break;
                case "劫財":  shiShenName = "劫財格:合伙真仗義"; break;
                case "食神":  shiShenName = "食神格:才藝真實，享受生活"; break;
                case "傷官":  shiShenName = "傷官格:創新真才華"; break;
            };
            return shiShenName;
        }

        // 輔助：計算十神名稱
        private static string GetShiShenName(char riGan, char targetGan)
        {
            string riWX = WX[riGan];
            string tgWX = WX[targetGan];
            bool sameYinYang = (WN[riGan] % 2) == (WN[targetGan] % 2); // 甲丙戊庚壬=陽

            if (tgWX == riWX) return sameYinYang ? "比肩" : "劫財";
            if (WuXingShengKe(riWX, tgWX) == "生") return sameYinYang ? "食神" : "傷官";  
            if (WuXingShengKe(riWX, tgWX) == "克") return sameYinYang ? "偏財" : "正財";
            if (WuXingShengKe(tgWX, riWX) == "生") return sameYinYang ? "偏印" : "正印";
            if (WuXingShengKe(tgWX, riWX) == "克") return sameYinYang ? "七殺" : "正官";  
            return "未知";
        }

        private static string WuXingShengKe(string a, string b)
        {
            string[,] matrix = {
            {"", "生", "克", "洩", "耗"},
            {"洩", "", "生", "克", "耗"},
            {"耗", "洩", "", "生", "克"},
            {"克", "耗", "洩", "", "生"},
            {"生", "克", "耗", "洩", ""}
            };
            string[] wx = { "木", "火", "土", "金", "水" };
            int i = Array.IndexOf(wx, a);
            int j = Array.IndexOf(wx, b);
            return matrix[i, j];
        }
        #region 輔助方法（請從您的「終極入口：三層合一分析.txt」複製所有方法）

        /// <summary>
        /// 根據日主和干支獲取十神名稱。
        /// </summary>
        private string GetShiShen(char dayMaster, char targetGan)
        {
            string riWX = WX[dayMaster];
            string tgWX = WX[targetGan];
            bool sameYinYang = (dayMaster % 2) == (targetGan % 2); // 甲丙戊庚壬=陽

            if (tgWX == riWX) return sameYinYang ? "比肩" : "劫財";
            if (WuXingShengKe(riWX, tgWX) == "生") return sameYinYang ? "正印" : "偏印";
            if (WuXingShengKe(riWX, tgWX) == "克") return sameYinYang ? "正官" : "七殺";
            if (WuXingShengKe(tgWX, riWX) == "生") return sameYinYang ? "食神" : "傷官";
            if (WuXingShengKe(tgWX, riWX) == "克") return sameYinYang ? "正財" : "偏財";
            return "未知";
        }

        /// <summary>
        /// 判斷某十神是否為真神（請貼上您的 IsTrue 邏輯）。
        /// </summary>
        private bool IsTrue(string shiShen)
        {
            // ...
            return true;
        }

        /// <summary>
        /// 判斷某十神是否為實神（請貼上您的 IsReal 邏輯）。
        /// </summary>
        private bool IsReal(string shiShen)
        {
            // ...
            return true;
        }

        // ... (請將文字檔中所有其他輔助方法複製貼上) ...

        #endregion

        /// <summary>
        /// 【核心入口】執行完整的盲派分析。
        /// </summary>
        /// <param name="yearPillar">年柱 (e.g., "甲子")</param>
        /// <param name="monthPillar">月柱</param>
        /// <param name="dayPillar">日柱</param>
        /// <param name="timePillar">時柱</param>
        /// <returns>BlindSchoolAnalysisResult 分析結果模型</returns>
        public BlindSchoolAnalysisResult Analyze(string yearPillar, string monthPillar, string dayPillar, string timePillar)
        {
            // 1. 提取日主 (日干)
            if (dayPillar.Length != 2)
            {
                throw new ArgumentException("日柱格式錯誤，必須是兩個字元（干支）。");
            }
            char riGan = dayPillar[0];
            string riWX = WX.ContainsKey(riGan) ? WX[riGan] : "未知";

            // 2. 組合八字字串
            string baziString = $"{yearPillar} {monthPillar} {dayPillar} {timePillar}";

            // 3. 執行分析邏輯
            // 在這裡呼叫您所有的判斷方法，填充 TrueFake, BodyUse, Pattern 等結構。
            // 由於我沒有您的完整文字檔，以下為結果填充的框架：

            var trueFakeInfo = new TrueFakeInfo();
            // 範例：分析年干的十神
            char yearGan = yearPillar[0];
            string yearShiShen = GetShiShen(riGan, yearGan);
            trueFakeInfo.ShiShens.Add(new ShiShenResult
            {
                Gan = yearGan,
                ShiShen = yearShiShen,
                WuXing = WX.ContainsKey(yearGan) ? WX[yearGan] : "",
                IsTrueGod = IsTrue(yearShiShen),
                IsReal = IsReal(yearShiShen)
            });
            // ... (重複此邏輯處理其他天干和地支藏干)

            var bodyUseInfo = new BodyUseInfo
            {
                CaiWX = "PlaceholderCaiWX",
                GuanWX = "PlaceholderGuanWX",
                // ... (根據您的邏輯填充其他欄位) ...
            };

            var patternInfo = new PatternInfo
            {
                Name = "PlaceholderPatternName",
                Description = "PlaceholderPatternDescription",
                // ...
            };


            // 4. 組織最終結果
            var analysisResult = new BlindSchoolAnalysisResult
            {
                BaZi = baziString,
                RiGan = riGan,
                RiWX = riWX,
                TrueFake = trueFakeInfo,
                BodyUse = bodyUseInfo,
                Pattern = patternInfo,
                Conclusion = $"日主 {riGan}({riWX}) 命理分析已完成。主要格局為 {patternInfo.Name}。"
            };

            return analysisResult;
        }
    }
}