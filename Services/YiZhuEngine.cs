using Ecanapi.Models;
using Ecanapi.Models.Ecanapi.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Ecanapi.Services
{
    public class YiZhuEngine
    {
        private static readonly string[] Stems = { "甲", "乙", "丙", "丁", "戊", "己", "庚", "辛", "壬", "癸" };
        private static readonly string[] Zhi = { "子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥" };

        public PillarAnalysisResult Diagnose(AstrologyChartResult data, int gender)
        {
            if (data?.Bazi?.DayPillar == null) return null;

            string stem = data.Bazi.DayPillar.HeavenlyStem;
            string branch = data.Bazi.DayPillar.EarthlyBranch;

            var result = new PillarAnalysisResult
            {
                DayPillar = stem + branch,
                DayMasterAnalysis = GetDayMasterText(stem)
            };

            // 1. 婚姻定數：自動抓取夫/妻星在日支的位階 (對應文稿 2.1)
            result.MarriageStatus = InferMarriageStatus(stem, branch, gender);

            // 2. 事業財富：自動抓取財星位階 (對應文稿 3.2)
            result.CareerStatus = InferCareerStatus(stem, branch);

            // 3. 子息產厄：自動抓取食傷位階 (對應文稿 2.2，包含辛亥壬水祿旺邏輯)
            result.ChildrenStatus = InferChildrenStatus(stem, branch,gender);

            // 4. 專家建議：結合能態與氣化感應
            result.RelativesAnalysis = InferExpertAdvice(stem, branch);

            return result;
        }

        private string InferMarriageStatus(string stem, string branch, int gender)
        {
            // 1. 定義十神對照
            // 男命看正財 (妻星), 女命看正官 (夫星)
            string targetStar = "";
            string starTitle = "";

            if (gender == 1) // 男命
            {
                starTitle = "妻星";
                targetStar = GetTenGodStar(stem, "正財");
            }
            else // 女命
            {
                starTitle = "夫星";
                targetStar = GetTenGodStar(stem, "正官");
            }

            // 2. 取得該星在「日支」的能量位 (例如：辛金男命，妻星甲木在亥位)
            string stage = GetLifeStage(targetStar, branch);

            // 3. 專家級特殊斷語 (針對辛亥日)
            if (stem == "辛" && branch == "亥")
            {
                if (gender == 1) // 男命辛亥
                    return $"· 婚姻定數：【日坐長生】妻星({targetStar})處亥位「長生」。主妻子端莊，得妻財或妻助，感情深厚。";
                else // 女命辛亥
                    return $"· 婚姻定數：【傷官坐命】夫星(丙火)處亥位「絕」地。感情易起波瀾，宜晚婚或修養心性。";
            }

            return $"· 婚姻定數：{starTitle}({targetStar})處日支「{stage}」位。{GetStageDescription(stage)}";
        }
        private string GetStageDescription(string stage)
        {
            return stage switch
            {
                "長生" => "屬生機勃勃、根基深厚之格，凡事多得貴人助。",
                "沐浴" => "主情感豐富，具備藝術才華，但需防情海浮沈。",
                "冠帶" => "象徵事業起步，具備領導潛力，名利雙收之象。",
                "臨官" => "代表祿旺之地，白手起家，具備極強競爭力。",
                "帝旺" => "能量達到巔峰，事業輝煌，但須注意剛愎自用。",
                "衰" => "氣勢趨緩，宜守成、不宜冒進，凡事求穩。",
                "病" => "能量稍弱，凡事不宜過度操勞，適合靜心修養。",
                "死" => "能量入庫，適合鑽研技術或宗教玄學，不宜爭鋒。",
                "墓" => "主守成、聚財，性格較內斂，適合幕後規劃。",
                "絕" => "屬置之死地而後生，大器晚成，宜修身養性。",
                "胎" => "孕育之象，具備無限可能，適合創業初期。",
                "養" => "培育之時，穩打穩紮，後勁十足。",
                _ => "能量平穩，順其自然發展即可。"
            };
        }

        // 輔助方法：根據日主取得對應的十神天干
        private string GetTenGodStar(string dayStem, string tenGodName)
        {
            var map = new Dictionary<string, Dictionary<string, string>> {
        { "甲", new Dictionary<string, string> { { "正財", "己" }, { "正官", "辛" }, { "傷官", "丁" } } },
        { "乙", new Dictionary<string, string> { { "正財", "戊" }, { "正官", "庚" }, { "傷官", "丙" } } },
        { "丙", new Dictionary<string, string> { { "正財", "辛" }, { "正官", "癸" }, { "傷官", "己" } } },
        { "丁", new Dictionary<string, string> { { "正財", "庚" }, { "正官", "壬" }, { "傷官", "戊" } } },
        { "戊", new Dictionary<string, string> { { "正財", "癸" }, { "正官", "乙" }, { "傷官", "辛" } } },
        { "己", new Dictionary<string, string> { { "正財", "壬" }, { "正官", "甲" }, { "傷官", "庚" } } },
        { "庚", new Dictionary<string, string> { { "正財", "乙" }, { "正官", "丁" }, { "傷官", "癸" } } },
        { "辛", new Dictionary<string, string> { { "正財", "甲" }, { "正官", "丙" }, { "傷官", "壬" } } },
        { "壬", new Dictionary<string, string> { { "正財", "丁" }, { "正官", "己" }, { "傷官", "乙" } } },
        { "癸", new Dictionary<string, string> { { "正財", "丙" }, { "正官", "戊" }, { "傷官", "甲" } } }
        };

            if (map.TryGetValue(dayStem, out var tenGods))
            {
                if (tenGods.TryGetValue(tenGodName, out var star)) return star;
            }
            return "";
        }

        // 輔助方法：根據日主取得對應的十神天干


        private string InferChildrenStatus(string stem, string branch, int gender)
        {
            // 1. 根據性別定義子息星 (男官殺、女食傷)
            // 這裡以辛金日主為例：男看丙/丁火，女看壬/癸水
            string childStar = "";
            string starName = "";

            if (gender == 1) // 男命
            {
                childStar = GetTenGodStar(stem, "正官"); // 這裡建議取官星
                starName = $"子息星({childStar}火)";
            }
            else // 女命
            {
                childStar = GetTenGodStar(stem, "傷官");
                starName = $"子息星({childStar}水)";
            }

            // 2. 取得能量位 (例如：壬水見亥為「臨官/祿旺」)
            string stage = GetLifeStage(childStar, branch);

            // 3. 根據能量位給予深度解說
            switch (stage)
            {
                case "長生":
                case "臨官":
                case "帝旺":
                    return $"· 子息產厄：【子女聰慧】{starName}處{branch}位「{stage}」。子女多才多藝，具備極強的競爭意識與創造力，未來必成大器。";

                case "冠帶":
                case "養":
                    return $"· 子息產厄：【後輩有為】{starName}處{branch}位「{stage}」。子女發展穩健，具備責任感，成年後能光耀門楣。";

                case "死":
                case "絕":
                    return $"· 子息產厄：{starName}處{branch}位「{stage}」。需注意首胎健康，子女發展較易受限於先天環境，宜多加耐心引導。";

                case "病":
                case "墓":
                    return $"· 子息產厄：{starName}處{branch}位「{stage}」。子女體質可能稍弱，但與父母緣分深厚，晚年有依。";

                default:
                    return $"· 子息產厄：{starName}處{branch}位「{stage}」。子息能量正常，晚年有依。";
            }
        }

        private string InferCareerStatus(string stem, string branch)
        {
            string moneyStar = GetTenGod(stem, "正財");
            string stage = GetLifeStage(moneyStar, branch);

            if (stem == "辛" && branch == "亥")
                return "· 事業財富：適合創意、設計、專業技術或靠技能獲利之行業。具備技術立身之定數。";

            return stage switch
            {
                "臨官" or "帝旺" => $"· 事業財富：財星坐「{stage}」祿位。具備極強的理財能量與經營天賦，財源厚實。",
                "墓" => "· 事業財富：財星入庫。主為人節儉守財，適合穩定儲蓄與不動產投資。",
                _ => $"· 事業財富：能量趨向「{stage}」。宜專業技術立身，穩健求財。"
            };
        }

        private string InferExpertAdvice(string stem, string branch)
        {
            if (stem == "辛" && branch == "亥")
                return "· 專家建議：【日坐食傷】才華溢出，感性豐富。具備極佳的表達力，宜強化執行力以貫徹志向。";

            return "· 專家建議：依據日柱能態，宜修身齊家，順應五行生剋規律，方能化解先天定數之不足。";
        }

        // --- 底層運算工具 ---

        private string GetTenGod(string dayStem, string target)
        {
            int idx = Array.IndexOf(Stems, dayStem);
            if (idx == -1) return "";
            return target switch
            {
                "正財" => Stems[((idx + 4) % 10) % 2 == 0 ? ((idx + 4) % 10) + 1 : ((idx + 4) % 10) - 1],
                "正官" => Stems[((idx + 6) % 10) % 2 == 0 ? ((idx + 6) % 10) + 1 : ((idx + 6) % 10) - 1],
                "傷官" => Stems[((idx + 2) % 10) % 2 == 0 ? ((idx + 2) % 10) + 1 : ((idx + 2) % 10) - 1],
                _ => dayStem
            };
        }

        private string GetLifeStage(string stem, string branch)
        {
            // 定義地支順序，用於計算位移
            string branches = "子丑寅卯辰巳午未申酉戌亥";
            int branchIdx = branches.IndexOf(branch);
            if (branchIdx == -1) return "未知";

            // 定義長生十二神名稱
            string[] stages = { "長生", "沐浴", "冠帶", "臨官", "帝旺", "衰", "病", "死", "墓", "絕", "胎", "養" };

            // 定義各天干「長生」起點的地支索引
            // 陽干：長生、沐浴...順行；陰干：長生、沐浴...逆行
            var startPos = new Dictionary<string, (int startIdx, bool isForward)>
    {
        { "甲", (2, true)  }, // 甲木長生在寅 (順) - 註：部分門派論亥，依通用標準取寅或亥
        { "丙", (5, true)  }, // 丙火長生在巳 (順) - 註：通用取寅，此處依標準長生表
        { "戊", (5, true)  }, // 戊土同丙火
        { "庚", (8, true)  }, // 庚金長生在申 (順) - 註：通用取巳
        { "壬", (11, true) }, // 壬水長生在亥 (順) - 註：通用取申
        
        // 依照經典《三命通會》標準長生表：
        { "乙", (6, false) }, // 乙木長生在午 (逆)
        { "丁", (9, false) }, // 丁火長生在酉 (逆)
        { "己", (9, false) }, // 己土同丁火
        { "辛", (0, false) }, // 辛金長生在子 (逆)
        { "癸", (3, false) }  // 癸水長生在卯 (逆)
    };

            // --- 專業對照表修正 (採用最通用的五行長生邏輯) ---
            var standardMap = new Dictionary<string, int> {
        {"甲", 11}, {"丙", 2}, {"戊", 2}, {"庚", 5}, {"壬", 8},  // 陽干長生位
        {"乙", 6},  {"丁", 9}, {"己", 9}, {"辛", 0}, {"癸", 3}   // 陰干長生位
    };

            if (!standardMap.ContainsKey(stem)) return "未知";

            int startIdx = standardMap[stem];
            bool isForward = "甲丙戊庚壬".Contains(stem);

            int offset;
            if (isForward)
            {
                offset = (branchIdx - startIdx + 12) % 12;
            }
            else
            {
                offset = (startIdx - branchIdx + 12) % 12;
            }

            return stages[offset];
        }

        private string GetDayMasterText(string stem)
        {
            return stem switch
            {
                "辛" => "【辛金】最清秀，為人處事多靈巧。創造力強且具競爭意識，外柔內剛，自尊心強。",
                "甲" => "【甲木】仁慈具開創力，領袖慾強，具備開拓精神。",
                _ => $"【{stem}】日元，具備五行基本特性。"
            };
        }
    }
}