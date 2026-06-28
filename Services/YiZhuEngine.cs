using Ecanapi.Models;
using Ecanapi.Models.Ecanapi.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ecanapi.Services
{
    /// <summary>
    /// 一柱論命引擎（六步完整算法）
    /// Step1: 有效日干（男女陰陽換位）
    /// Step2: 虛辰遁法定六親地支位
    /// Step3: 六親十二長生強弱 + 日支刑沖害合
    /// Step4: 十神非六親命理象意
    /// Step5: 月令對日干喜忌 → 先天根基
    /// Step6: 旬中干支 vs 日柱互動 × 喜忌
    /// </summary>
    public class YiZhuEngine
    {
        private static readonly string[] Stems =
            { "甲", "乙", "丙", "丁", "戊", "己", "庚", "辛", "壬", "癸" };

        private static readonly string[] Branches =
            { "子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥" };

        private static readonly string[] LifeStages =
            { "長生", "沐浴", "冠帶", "臨官", "帝旺", "衰", "病", "死", "墓", "絕", "胎", "養" };

        // 各天干長生起點（Branches 索引）
        private static readonly Dictionary<string, int> LifeStageStart = new()
        {
            { "甲", 11 }, { "丙", 2 }, { "戊", 2 }, { "庚", 5 }, { "壬", 8 },
            { "乙", 6  }, { "丁", 9 }, { "己", 9 }, { "辛", 0 }, { "癸", 3 }
        };

        // 五行索引：0=木 1=火 2=土 3=金 4=水
        private static readonly int[] StemElement = { 0, 0, 1, 1, 2, 2, 3, 3, 4, 4 };

        private static readonly Dictionary<string, int> BranchMonthElement = new()
        {
            { "寅", 0 }, { "卯", 0 },
            { "巳", 1 }, { "午", 1 },
            { "辰", 2 }, { "戌", 2 }, { "丑", 2 }, { "未", 2 },
            { "申", 3 }, { "酉", 3 },
            { "亥", 4 }, { "子", 4 }
        };

        // 地支六合
        private static readonly HashSet<string> LiuHeSet = new()
        {
            "子丑", "丑子", "寅亥", "亥寅", "卯戌", "戌卯",
            "辰酉", "酉辰", "巳申", "申巳", "午未", "未午"
        };

        // 地支六沖
        private static readonly HashSet<string> LiuChongSet = new()
        {
            "子午", "午子", "丑未", "未丑", "寅申", "申寅",
            "卯酉", "酉卯", "辰戌", "戌辰", "巳亥", "亥巳"
        };

        // 地支六害
        private static readonly HashSet<string> LiuHaiSet = new()
        {
            "子未", "未子", "丑午", "午丑", "寅巳", "巳寅",
            "卯辰", "辰卯", "申亥", "亥申", "酉戌", "戌酉"
        };

        // 地支三合（各組）
        private static readonly List<HashSet<string>> SanHeGroups = new()
        {
            new() { "申", "子", "辰" },
            new() { "寅", "午", "戌" },
            new() { "巳", "酉", "丑" },
            new() { "亥", "卯", "未" }
        };

        // 地支三刑
        private static readonly HashSet<string> SanXingSet = new()
        {
            "寅巳", "巳寅", "巳申", "申巳", "申寅", "寅申",
            "丑戌", "戌丑", "戌未", "未戌", "未丑", "丑未",
            "子卯", "卯子"
        };

        // 天干五合
        private static readonly Dictionary<string, string> StemHeMap = new()
        {
            { "甲", "己" }, { "己", "甲" }, { "乙", "庚" }, { "庚", "乙" },
            { "丙", "辛" }, { "辛", "丙" }, { "丁", "壬" }, { "壬", "丁" },
            { "戊", "癸" }, { "癸", "戊" }
        };

        // 天干相沖（甲庚 乙辛 丙壬 丁癸）
        private static readonly HashSet<string> StemChongSet = new()
        {
            "甲庚", "庚甲", "乙辛", "辛乙", "丙壬", "壬丙", "丁癸", "癸丁"
        };

        // ==========================================
        // 主入口：六步完整分析
        // ==========================================

        /// <summary>
        /// 一柱論命六步完整分析
        /// </summary>
        /// <param name="yongShenElem">八字格局用神五行（如"水"），若提供則覆蓋月令喜忌判定</param>
        public string Analyze(string dayStem, string dayBranch, string monthBranch, int gender,
            string? yongShenElem = null)
        {
            // Step 1: 有效日干
            string effStem = GetEffectiveStem(dayStem, gender);

            // Step 2: 旬首 + 旬中地支映射 + 空亡
            string xunShou = GetXunShou(dayStem, dayBranch);
            var xunMap = GetXunMapping(xunShou);           // stem -> branch
            var kongWang = GetKongWang(xunShou);           // 空亡地支列表

            // Step 5 先算（Step 6 需要喜忌資訊）
            var (bodyStrength, xiElements, jiElements) = GetBodyStrength(dayStem, monthBranch);

            // 格局用神覆蓋：若外部傳入 yongShenElem，以此為準判定六親喜忌標注
            string[]? xiOverride = null;
            var elemMap = new Dictionary<string, int> { {"木",0},{"火",1},{"土",2},{"金",3},{"水",4} };
            if (!string.IsNullOrEmpty(yongShenElem) && elemMap.TryGetValue(yongShenElem, out int yel))
            {
                // 喜：用神 + 生用神的五行；忌：其餘
                xiOverride = new[] { yongShenElem, Fe((yel + 4) % 5) };
            }

            var sb = new StringBuilder();
            sb.AppendLine($"【六親定數 · {dayStem}{dayBranch}日（{(gender == 1 ? "男" : "女")}命）】");
            sb.AppendLine($"有效干：{dayStem}　旬首：{xunShou}　空亡：{string.Join("、", kongWang)}");
            sb.AppendLine();

            // 驗證表：日干十神六親對照
            sb.AppendLine("▍日干十神六親對照");
            var verifyParts = new List<string>();
            foreach (string s in Stems)
            {
                string tg  = GetTenGodName(dayStem, s);
                string rel = (s == dayStem) ? "本命" : GetRelativeName(tg, gender);
                verifyParts.Add($"{s}={tg}({rel})");
            }
            sb.AppendLine(string.Join("　", verifyParts.Take(5)));
            sb.AppendLine(string.Join("　", verifyParts.Skip(5)));
            sb.AppendLine();

            // Step 3 + 4
            sb.AppendLine("▍六親定位與命理象意");
            foreach (string stem in Stems)
            {
                // 問題3修正：跳過日柱本身（日干在旬中坐日支，即命主本位，非六親）
                if (stem == dayStem) continue;

                string branch  = xunMap[stem];
                string tenGod  = GetTenGodName(dayStem, stem);
                string relative = GetRelativeName(tenGod, gender);
                string stage    = GetLifeStage(stem, branch);
                string dayRel   = GetBranchRelation(dayBranch, branch);
                bool isKong     = kongWang.Contains(branch);
                // 喜忌判斷：優先用格局用神覆蓋值，否則用月令計算值
                int stemEl = StemElement[Array.IndexOf(Stems, stem)];
                string stemFe = Fe(stemEl);
                bool isXi = (xiOverride != null) ? xiOverride.Contains(stemFe) : xiElements.Contains(stemFe);
                string nonRelMeaning = GetTenGodNonRelMeaning(tenGod, stage, isXi);

                sb.AppendLine(BuildRelativeLine(relative, tenGod, stem, branch, stage, dayRel, isKong, nonRelMeaning));
            }
            sb.AppendLine();

            // Step 5
            sb.AppendLine("▍先天根基（月令喜忌）");
            sb.AppendLine(BuildBodyStrengthText(dayStem, monthBranch, bodyStrength, xiElements, jiElements, yongShenElem));
            sb.AppendLine();

            // Step 6（喜忌標注同樣優先用格局用神）
            sb.AppendLine("▍旬中干支互動");
            var xiForStep6 = xiOverride ?? xiElements;
            sb.Append(BuildXunInteractions(dayStem, dayBranch, dayStem, xunMap, xiForStep6, gender));

            return sb.ToString().Trim();
        }

        // ==========================================
        // Step 1：有效日干
        // ==========================================

        private static string GetEffectiveStem(string dayStem, int gender)
        {
            bool isYang = "甲丙戊庚壬".Contains(dayStem);
            bool needSwitch = (gender == 1 && !isYang) || (gender == 0 && isYang);
            if (!needSwitch) return dayStem;

            var pair = new Dictionary<string, string>
            {
                { "甲","乙" }, { "乙","甲" }, { "丙","丁" }, { "丁","丙" }, { "戊","己" },
                { "己","戊" }, { "庚","辛" }, { "辛","庚" }, { "壬","癸" }, { "癸","壬" }
            };
            return pair[dayStem];
        }

        // ==========================================
        // Step 2：虛辰遁法
        // ==========================================

        private static string GetXunShou(string dayStem, string dayBranch)
        {
            int si = Array.IndexOf(Stems, dayStem);
            int bi = Array.IndexOf(Branches, dayBranch);
            int startBi = (bi - si % 12 + 12) % 12;
            return "甲" + Branches[startBi];
        }

        private static Dictionary<string, string> GetXunMapping(string xunShou)
        {
            int start = Array.IndexOf(Branches, xunShou[1].ToString());
            var map = new Dictionary<string, string>();
            for (int i = 0; i < 10; i++)
                map[Stems[i]] = Branches[(start + i) % 12];
            return map;
        }

        private static List<string> GetKongWang(string xunShou)
        {
            int start = Array.IndexOf(Branches, xunShou[1].ToString());
            var used = new HashSet<int>(Enumerable.Range(0, 10).Select(i => (start + i) % 12));
            return Branches.Where((_, i) => !used.Contains(i)).ToList();
        }

        // ==========================================
        // 十神計算
        // ==========================================

        private static string GetTenGodName(string dayStem, string targetStem)
        {
            if (dayStem == targetStem) return "比肩";

            int di = Array.IndexOf(Stems, dayStem);
            int ti = Array.IndexOf(Stems, targetStem);
            int de = StemElement[di];
            int te = StemElement[ti];
            bool sameP = (di % 2) == (ti % 2);

            if (de == te) return sameP ? "比肩" : "劫財";
            if ((de + 1) % 5 == te) return sameP ? "食神" : "傷官";   // 我生
            if ((de + 2) % 5 == te) return sameP ? "偏財" : "正財";   // 我剋
            if ((te + 2) % 5 == de) return sameP ? "七殺" : "正官";   // 剋我
            if ((te + 1) % 5 == de) return sameP ? "偏印" : "正印";   // 生我
            return "比肩";
        }

        private static string GetRelativeName(string tenGod, int gender)
        {
            // effStem 設計：男命一律用陽干計算（比肩=陽=兄弟，劫財=陰=姊妹）
            //               女命一律用陰干計算（比肩=陰=姊妹，劫財=陽=兄弟）
            //               男命子女來自官殺（七殺=陽官=兒，正官=陰官=女）
            //               女命子女來自食傷（傷官=陽食=兒，食神=陰食=女）
            if (gender == 1) // 男命（effStem 為陽干）
                return tenGod switch
                {
                    "比肩" => "兄弟", "劫財" => "姊妹",
                    "食神" => "女婿", "傷官" => "媳婦",
                    "偏財" => "父",   "正財" => "妻",
                    "七殺" => "兒子", "正官" => "女兒",
                    "偏印" => "祖父", "正印" => "母",
                    _ => tenGod
                };
            else // 女命（effStem 為陰干）
                return tenGod switch
                {
                    "比肩" => "姊妹", "劫財" => "兄弟",
                    "食神" => "女兒", "傷官" => "兒子",
                    "偏財" => "繼父", "正財" => "父",
                    "七殺" => "夫之兄", "正官" => "夫",
                    "偏印" => "母",   "正印" => "祖母",
                    _ => tenGod
                };
        }

        // ==========================================
        // Step 3：地支關係 + 六親強弱
        // ==========================================

        public static string GetBranchRelation(string b1, string b2)
        {
            if (b1 == b2) return "";
            string key = b1 + b2;
            if (LiuHeSet.Contains(key))    return "六合";
            if (LiuChongSet.Contains(key)) return "相沖";
            if (LiuHaiSet.Contains(key))   return "相害";
            if (SanXingSet.Contains(key))  return "相刑";
            // 問題1修正：兩支相遇只是半合，非三合（三合需三支同聚）
            foreach (var g in SanHeGroups)
                if (g.Contains(b1) && g.Contains(b2)) return "半合";
            return "";
        }

        private static string ClassifyStrength(string stage) => stage switch
        {
            "長生" or "冠帶" or "臨官" or "帝旺" => "旺",
            "沐浴" or "養" => "中",
            "衰" or "病" => "弱",
            _ => "極弱"   // 死 墓 絕 胎
        };

        private static string BuildRelativeLine(
            string relative, string tenGod, string stem, string branch,
            string stage, string dayRel, bool isKong, string nonRelMeaning)
        {
            string strength = ClassifyStrength(stage);
            string kongNote = isKong ? "【空亡】" : "";

            // 六親關係斷語
            string relDesc = (strength, dayRel) switch
            {
                ("旺",  "相沖") => "六親有力，然相沖主分離聚少離多",
                ("旺",  "六合") => "六親有力，相合情深，多得助力",
                ("旺",  "相害") => "六親有力，然暗中有損，需防隱患",
                ("旺",  "半合") => "六親有力，半合拱局，緣份深厚",
                ("旺",  "相刑") => "六親有力，然相刑主紛爭口舌",
                ("旺",  _)      => "六親有力，先天根基穩固",
                ("極弱","相沖") => "六親緣薄，且相沖，刑剋最重",
                ("極弱","六合") => "六親雖弱，相合有情，緣份勉強維繫",
                ("極弱",_)      => "六親緣薄，先天刑剋，緣分不深",
                ("弱",  "相沖") => "六親稍弱，且相沖，聚少離多",
                ("弱",  _)      => "六親稍弱，宜後天努力維繫",
                (_,     "相沖") => "有情有緣，然相沖主波折",
                (_,     "六合") => "相合情融，六親緣份佳",
                (_,     _)      => "六親緣份平常"
            };

            if (isKong) relDesc += "；逢空亡，需填實方顯";

            string dayRelNote = !string.IsNullOrEmpty(dayRel) ? $"，日支{dayRel}" : "";

            // 非六親象意
            string nonRelNote = !string.IsNullOrEmpty(nonRelMeaning) ? $"　{nonRelMeaning}" : "";

            return $"· {relative}（{tenGod}{stem} · {branch}{kongNote} · {stage}）{dayRelNote}：{relDesc}。{nonRelNote}";
        }

        // ==========================================
        // Step 4：十神非六親象意
        // ==========================================

        private static string GetTenGodNonRelMeaning(string tenGod, string stage, bool isXi)
        {
            string strength = ClassifyStrength(stage);
            bool isStrong = strength is "旺" or "中";

            if (isXi)
                return tenGod switch
                {
                    "正財" => isStrong ? "【財運】財源穩定，薪資豐厚。"       : "【財運】正財薄弱，錢財難聚。",
                    "偏財" => isStrong ? "【偏財】橫財機會多，理財靈活。"     : "【偏財】橫財難得，偏業無緣。",
                    "正官" => isStrong ? "【官職】仕途順遂，名譽佳。"         : "【官職】官職難求，名譽易損。",
                    "七殺" => isStrong ? "【競爭】競爭力強，衝勁十足。"       : "【競爭】壓力雖重，能化為衝勁。",
                    "正印" => isStrong ? "【貴人】文書順利，貴人相助。"       : "【貴人】貴人難求，學業不利。",
                    "偏印" => isStrong ? "【技藝】術業有專攻，宗教緣深。"     : "【技藝】偏業難成，心神不定。",
                    "食神" => isStrong ? "【才藝】口福佳，才藝出眾，壽元足。" : "【才藝】才藝平平，口福稍薄。",
                    "傷官" => isStrong ? "【才華】才華橫溢，口才極佳。"       : "【才華】才華受阻，官運不順。",
                    "比肩" => isStrong ? "【朋友】平輩助力多，合夥有利。"     : "【朋友】同輩競爭，合夥需謹慎。",
                    "劫財" => isStrong ? "【助力】同氣幫身，鬥志增強。"       : "【助力】比劫雖弱，仍可扶身。",
                    _ => ""
                };
            else
                return tenGod switch
                {
                    "正財" => isStrong ? "【財運】財雖有根，然屬忌神，財來財去，難以積聚。" : "【財運】財星忌神且弱，破財耗財之象。",
                    "偏財" => isStrong ? "【偏財】偏財屬忌，橫財易得易失，散財之象。"       : "【偏財】偏財忌神且弱，投機損耗。",
                    "正官" => isStrong ? "【官職】官殺屬忌，名位雖有，壓力沉重難承。"       : "【官職】官星忌神且弱，仕途坎坷。",
                    "七殺" => isStrong ? "【競爭】七殺屬忌，外部壓力極重，易招官非爭端。"   : "【競爭】七殺忌神且弱，紛爭雖輕，仍宜謹慎。",
                    "正印" => isStrong ? "【貴人】印星屬忌，依賴心重，文書反成束縛。"       : "【貴人】印星忌神且弱，貴人無力。",
                    "偏印" => isStrong ? "【技藝】梟印屬忌，鑽牛角尖，思路易偏執。"         : "【技藝】梟印忌神且弱，雜學無成。",
                    "食神" => isStrong ? "【才藝】食神屬忌，洩氣過重，雖有才藝卻耗損精神。" : "【才藝】食傷忌神且弱，才藝有限。",
                    "傷官" => isStrong ? "【才華】傷官屬忌，才高易惹非議，官運受損。"       : "【才華】傷官忌神且弱，口才受限。",
                    "比肩" => isStrong ? "【朋友】比肩屬忌（身旺），同輩爭財，合夥損耗。"   : "【朋友】比肩忌神且弱，同儕拖累有限。",
                    "劫財" => isStrong ? "【破財】劫財屬忌，破財耗損最烈，防競爭損失。"     : "【破財】劫財忌神且弱，破財較輕。",
                    _ => ""
                };
        }

        // ==========================================
        // Step 5：月令喜忌 → 先天根基
        // ==========================================

        private static (string strength, string[] xiElements, string[] jiElements)
            GetBodyStrength(string dayStem, string monthBranch)
        {
            int dayEl = StemElement[Array.IndexOf(Stems, dayStem)];
            int monthEl = BranchMonthElement.TryGetValue(monthBranch, out int me) ? me : 2;

            string rel;
            if (monthEl == dayEl)                   rel = "比劫";
            else if ((monthEl + 1) % 5 == dayEl)   rel = "印";      // 月令生日主
            else if ((dayEl + 1) % 5 == monthEl)   rel = "食傷";   // 日主生月令
            else if ((dayEl + 2) % 5 == monthEl)   rel = "財";     // 日主剋月令
            else                                     rel = "官殺";   // 月令剋日主

            string strength = rel switch
            {
                "比劫"  => "身旺",
                "印"    => "偏強",
                "食傷"  => "偏弱",
                "財"    => "偏弱",
                "官殺"  => "身弱",
                _       => "中和"
            };

            // 喜用五行
            string[] xiElements = (strength is "身旺" or "偏強")
                ? new[] { Fe((dayEl + 1) % 5), Fe((dayEl + 2) % 5), Fe((dayEl + 3) % 5) }
                : new[] { Fe((dayEl + 4) % 5), Fe(dayEl) };

            string[] jiElements = (strength is "身旺" or "偏強")
                ? new[] { Fe((dayEl + 4) % 5), Fe(dayEl) }
                : new[] { Fe((dayEl + 1) % 5), Fe((dayEl + 2) % 5), Fe((dayEl + 3) % 5) };

            return (strength, xiElements, jiElements);
        }

        private static string Fe(int el) => el switch
        {
            0 => "木", 1 => "火", 2 => "土", 3 => "金", 4 => "水", _ => ""
        };

        private static string BuildBodyStrengthText(
            string dayStem, string monthBranch, string bodyStrength,
            string[] xiElements, string[] jiElements, string? yongShenElem = null)
        {
            int dayEl = StemElement[Array.IndexOf(Stems, dayStem)];
            int monthEl = BranchMonthElement.TryGetValue(monthBranch, out int me) ? me : 2;
            var sb = new StringBuilder();
            sb.AppendLine($"生於{monthBranch}月（{Fe(monthEl)}旺），日主{dayStem}{Fe(dayEl)}，{bodyStrength}。");
            sb.AppendLine($"  月令基礎：喜{string.Join("、", xiElements)}；忌{string.Join("、", jiElements)}。");
            if (!string.IsNullOrEmpty(yongShenElem))
                sb.AppendLine($"  ▷ 格局用神：{yongShenElem}（六親喜忌標注以格局用神為準）");
            return sb.ToString().TrimEnd();
        }

        // ==========================================
        // Step 6：旬中干支 vs 日柱互動
        // ==========================================

        private static string BuildXunInteractions(
            string dayStem, string dayBranch, string effStem,
            Dictionary<string, string> xunMap, string[] xiElements, int gender)
        {
            var sb = new StringBuilder();

            foreach (string stem in Stems)
            {
                if (stem == dayStem) continue;
                string branch  = xunMap[stem];
                if (branch == dayBranch) continue;

                string stemRel   = GetStemRelation(dayStem, stem);
                string branchRel = GetBranchRelation(dayBranch, branch);
                if (string.IsNullOrEmpty(stemRel) && string.IsNullOrEmpty(branchRel)) continue;

                string tenGod   = GetTenGodName(dayStem, stem);
                string relative = GetRelativeName(tenGod, gender);
                int stemEl      = StemElement[Array.IndexOf(Stems, stem)];
                bool isXi       = xiElements.Contains(Fe(stemEl));
                string xiJi     = isXi ? "喜" : "忌";

                bool bothActive = !string.IsNullOrEmpty(stemRel) && !string.IsNullOrEmpty(branchRel);
                string intensity = bothActive ? "天地俱動，影響最烈" : "";

                string impact = (isXi, branchRel, stemRel) switch
                {
                    (true,  "六合",  _)      => "喜神相合，吉上加吉",
                    (true,  "半合",  _)      => "喜神半合拱局，增強吉象",
                    (true,  "相沖",  "相沖") => "喜神天地俱沖，吉中帶憂",
                    (true,  "相沖",  _)      => "喜神被地沖，有所折損",
                    (true,  "相害",  _)      => "喜神受害，暗中有損",
                    (true,  _,       "相合") => "喜神天合，有情有義",
                    (true,  _,       _)      => "喜神有動，裨益命主",
                    (false, "相沖",  "相沖") => "忌神天地俱沖，凶象最烈",
                    (false, "相沖",  _)      => "忌神沖日，凶象較重",
                    (false, "六合",  _)      => "忌神相合，暗損難防",
                    (false, "相害",  _)      => "忌神相害，隱患纏身",
                    (false, "相刑",  _)      => "忌神相刑，紛爭官非",
                    (false, _,       "相合") => "忌神天合，需防被拖累",
                    _                        => isXi ? "喜神互動，有所助益" : "忌神互動，需加防範"
                };

                var parts = new List<string>();
                if (!string.IsNullOrEmpty(stemRel))   parts.Add($"天干{stemRel}（{dayStem}{stem}）");
                if (!string.IsNullOrEmpty(branchRel)) parts.Add($"地支{branchRel}（{dayBranch}{branch}）");
                if (!string.IsNullOrEmpty(intensity)) parts.Add(intensity);

                sb.AppendLine($"· {stem}{branch}（{relative}·{tenGod}·{xiJi}神）：{string.Join("，", parts)}。{impact}。");
            }

            return sb.ToString().TrimEnd();
        }

        private static string GetStemRelation(string s1, string s2)
        {
            if (StemHeMap.TryGetValue(s1, out var he) && he == s2) return "相合";
            if (StemChongSet.Contains(s1 + s2)) return "相沖";
            return "";
        }

        // ==========================================
        // 公用：十二長生
        // ==========================================

        public string GetLifeStage(string stem, string branch)
        {
            int bi = Array.IndexOf(Branches, branch);
            if (bi == -1 || !LifeStageStart.TryGetValue(stem, out int si)) return "未知";
            bool forward = "甲丙戊庚壬".Contains(stem);
            int offset = forward ? (bi - si + 12) % 12 : (si - bi + 12) % 12;
            return LifeStages[offset];
        }

        // ==========================================
        // 舊版 Diagnose（向下相容保留）
        // ==========================================

        public PillarAnalysisResult Diagnose(AstrologyChartResult data, int gender)
        {
            if (data?.Bazi?.DayPillar == null) return null;

            string stem   = data.Bazi.DayPillar.HeavenlyStem;
            string branch = data.Bazi.DayPillar.EarthlyBranch;

            return new PillarAnalysisResult
            {
                DayPillar         = stem + branch,
                DayMasterAnalysis = GetDayMasterText(stem),
                MarriageStatus    = InferMarriageStatus(stem, branch, gender),
                CareerStatus      = InferCareerStatus(stem, branch),
                ChildrenStatus    = InferChildrenStatus(stem, branch, gender),
                RelativesAnalysis = InferExpertAdvice(stem, branch)
            };
        }

        private string InferMarriageStatus(string stem, string branch, int gender)
        {
            string effStem = GetEffectiveStem(stem, gender);
            string spouseTenGod = gender == 1 ? "正財" : "正官";
            string spouseStem = GetTenGodStemByName(effStem, spouseTenGod);
            string stage = GetLifeStage(spouseStem, branch);
            string title = gender == 1 ? "妻星" : "夫星";
            return $"· 婚姻定數：{title}（{spouseStem}）處日支「{stage}」。{GetStageDesc(stage)}";
        }

        private string InferCareerStatus(string stem, string branch)
        {
            string moneyStem = GetTenGodStemByName(stem, "正財");
            string stage = GetLifeStage(moneyStem, branch);
            return stage switch
            {
                "臨官" or "帝旺" => $"· 事業財富：財星坐「{stage}」，財源豐厚，經營天賦強。",
                "墓"             => "· 事業財富：財星入庫，節儉守財，適合穩健投資。",
                _                => $"· 事業財富：財星「{stage}」，宜專業技術立身，穩健求財。"
            };
        }

        private string InferChildrenStatus(string stem, string branch, int gender)
        {
            string effStem = GetEffectiveStem(stem, gender);
            string childTenGod = gender == 1 ? "七殺" : "食神";
            string childStem = GetTenGodStemByName(effStem, childTenGod);
            string stage = GetLifeStage(childStem, branch);
            return stage switch
            {
                "長生" or "臨官" or "帝旺" => $"· 子息定數：子女聰慧（{childStem}·{stage}），有力旺相，未來必成大器。",
                "死" or "絕"               => $"· 子息定數：子女緣薄（{childStem}·{stage}），需耐心引導。",
                _                          => $"· 子息定數：子女緣份平常（{childStem}·{stage}），晚年有依。"
            };
        }

        private string InferExpertAdvice(string stem, string branch)
            => $"· 建議：依據{stem}{branch}日柱能態，宜修身齊家，順應五行規律。";

        private string GetDayMasterText(string stem)
            => $"【{stem}】日元，{Fe(StemElement[Array.IndexOf(Stems, stem)])}性命主，具備相應五行特質。";

        private static string GetStageDesc(string stage) => stage switch
        {
            "長生" => "根基深厚，感情穩定。",
            "沐浴" => "情感豐富，具藝術才華。",
            "冠帶" => "事業起步，名利雙收。",
            "臨官" => "祿旺之地，白手起家。",
            "帝旺" => "巔峰能量，感情強勢。",
            "衰"   => "氣勢趨緩，宜守成。",
            "病"   => "能量稍弱，宜靜心。",
            "死"   => "能量入庫，適合鑽研。",
            "墓"   => "守成聚財，性格內斂。",
            "絕"   => "大器晚成，宜修養。",
            "胎"   => "孕育之象，可能性多。",
            "養"   => "後勁十足，穩打穩紮。",
            _      => "順其自然。"
        };

        // 根據十神名稱取對應天干
        private static string GetTenGodStemByName(string dayStem, string tenGodName)
        {
            foreach (string target in Stems)
            {
                if (GetTenGodName(dayStem, target) == tenGodName)
                    return target;
            }
            return dayStem;
        }
    }
}
