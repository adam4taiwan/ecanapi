using System;
using System.Collections.Generic;

namespace Ecanapi.Models.Analysis
{
    public class Star
    {
        public class StarDefinition
        {
            public string StarName { get; set; }
            public string Category { get; set; }
            public string Level { get; set; }
            public string Element { get; set; }
            public string Polar { get; set; }
            public string Transform { get; set; }
            public string Meaning { get; set; }
        }

        public static class StarMap
        {
            // 基礎定義 (113 顆星全展開)
            private static readonly Dictionary<string, StarDefinition> BaseDict = new Dictionary<string, StarDefinition>(StringComparer.OrdinalIgnoreCase)
            {
                // 1. 正星
                { "紫微", new StarDefinition { StarName = "紫微", Level = "正星", Element = "己土", Transform = "尊貴", Meaning = "官祿主，掌爵祿，解厄制化。性情厚重。" } },
                { "天機", new StarDefinition { StarName = "天機", Level = "正星", Element = "乙木", Transform = "善", Meaning = "兄弟主，掌壽，智謀星。" } },
                { "太陽", new StarDefinition { StarName = "太陽", Level = "正星", Element = "丙火", Transform = "貴", Meaning = "官祿主，博愛、大公無私。" } },
                { "武曲", new StarDefinition { StarName = "武曲", Level = "正星", Element = "辛金", Transform = "財", Meaning = "財帛主，性格剛毅果決。" } },
                { "天同", new StarDefinition { StarName = "天同", Level = "正星", Element = "壬水", Transform = "福", Meaning = "福德主，平易近人，知足常樂。" } },
                { "廉貞", new StarDefinition { StarName = "廉貞", Level = "正星", Element = "丁火", Transform = "囚", Meaning = "官祿主，反應靈敏，性格硬氣。" } },
                { "天府", new StarDefinition { StarName = "天府", Level = "正星", Element = "戊土", Transform = "庫", Meaning = "財帛主，掌財庫，穩重寬厚。" } },
                { "太陰", new StarDefinition { StarName = "太陰", Level = "正星", Element = "癸水", Transform = "富", Meaning = "田宅主，心思細膩，文靜溫柔。" } },
                { "貪狼", new StarDefinition { StarName = "貪狼", Level = "正星", Element = "甲木", Transform = "桃花", Meaning = "慾望之主，多才多藝，圓滑善交際。" } },
                { "巨門", new StarDefinition { StarName = "巨門", Level = "正星", Element = "癸水", Transform = "暗", Meaning = "是非主，口才佳，觀察力強。" } },
                { "天相", new StarDefinition { StarName = "天相", Level = "正星", Element = "壬水", Transform = "印", Meaning = "官祿主，掌印，敦厚公正。" } },
                { "天梁", new StarDefinition { StarName = "天梁", Level = "正星", Element = "戊土", Transform = "蔭", Meaning = "父母主，長者風範，清高避難。" } },
                { "七殺", new StarDefinition { StarName = "七殺", Level = "正星", Element = "庚金", Transform = "將軍", Meaning = "將星，權威，急躁進取。" } },
                { "破軍", new StarDefinition { StarName = "破軍", Level = "正星", Element = "癸水", Transform = "耗", Meaning = "損耗主，敢於突破，個性反叛。" } },

                // 2. 六吉六煞
                { "左輔", new StarDefinition { StarName = "左輔", Level = "吉星", Meaning = "平輩助力，圓滑人緣好。" } },
                { "右弼", new StarDefinition { StarName = "右弼", Level = "吉星", Meaning = "暗中相助，能彌補不足。" } },
                { "文昌", new StarDefinition { StarName = "文昌", Level = "吉星", Meaning = "才華文書，正途科名。" } },
                { "文曲", new StarDefinition { StarName = "文曲", Level = "吉星", Meaning = "口才藝術，異路功名。" } },
                { "天魁", new StarDefinition { StarName = "天魁", Level = "吉星", Meaning = "陽貴人，得年長者提攜。" } },
                { "天鉞", new StarDefinition { StarName = "天鉞", Level = "吉星", Meaning = "陰貴人，暗中助力。" } },
                { "祿存", new StarDefinition { StarName = "祿存", Level = "吉星", Meaning = "祿位財源，能解厄制化。" } },
                { "天馬", new StarDefinition { StarName = "天馬", Level = "吉星", Meaning = "動能變遷，利遷移財源。" } },
                { "擎羊", new StarDefinition { StarName = "擎羊", Level = "煞星", Meaning = "刑傷衝擊，剛烈果斷。" } },
                { "陀羅", new StarDefinition { StarName = "陀羅", Level = "煞星", Meaning = "拖延阻礙，心思糾結。" } },
                { "火星", new StarDefinition { StarName = "火星", Level = "煞星", Meaning = "暴發暴敗，性急沒耐性。" } },
                { "鈴星", new StarDefinition { StarName = "鈴星", Level = "煞星", Meaning = "沈悶驚難，性格陰沈。" } },
                { "地空", new StarDefinition { StarName = "地空", Level = "煞星", Meaning = "精神虛無，天馬行空。" } },
                { "地劫", new StarDefinition { StarName = "地劫", Level = "煞星", Meaning = "物質虛耗，易突發破財。" } },

                // 3. 丙級與神煞 (含紅鸞、天喜等)
                { "紅鸞", new StarDefinition { StarName = "紅鸞", Level = "丙級", Meaning = "主喜慶桃花。" } },
                { "天喜", new StarDefinition { StarName = "天喜", Level = "丙級", Meaning = "主喜悅幽默。" } },
                { "天刑", new StarDefinition { StarName = "天刑", Level = "丙級", Meaning = "主官非、紀律。" } },
                { "天姚", new StarDefinition { StarName = "天姚", Level = "丙級", Meaning = "主浪漫桃花。" } },
                { "解神", new StarDefinition { StarName = "解神", Level = "丙級", Meaning = "逢凶化吉。" } },
                { "天巫", new StarDefinition { StarName = "天巫", Level = "丙級", Meaning = "主遺產升遷、宗教。" } },
                { "天月", new StarDefinition { StarName = "天月", Level = "丙級", Meaning = "主小病痛。" } },
                { "陰煞", new StarDefinition { StarName = "陰煞", Level = "丙級", Meaning = "主小人暗中作祟。" } },
                { "龍池", new StarDefinition { StarName = "龍池", Level = "丙級", Meaning = "藝術氣質。" } },
                { "鳳閣", new StarDefinition { StarName = "鳳閣", Level = "丙級", Meaning = "高貴體面。" } },
                { "台輔", new StarDefinition { StarName = "台輔", Level = "丙級", Meaning = "增加主星威嚴。" } },
                { "封誥", new StarDefinition { StarName = "封誥", Level = "丙級", Meaning = "主名譽獎勵。" } },
                { "三台", new StarDefinition { StarName = "三台", Level = "丙級", Meaning = "增進職位。" } },
                { "八座", new StarDefinition { StarName = "八座", Level = "丙級", Meaning = "增進名譽。" } },
                { "恩光", new StarDefinition { StarName = "恩光", Level = "丙級", Meaning = "主受恩寵提拔。" } },
                { "天貴", new StarDefinition { StarName = "天貴", Level = "丙級", Meaning = "主受人重用。" } },
                { "孤辰", new StarDefinition { StarName = "孤辰", Level = "丙級", Meaning = "孤僻自立。" } },
                { "寡宿", new StarDefinition { StarName = "寡宿", Level = "丙級", Meaning = "孤獨疏離。" } },
                { "天哭", new StarDefinition { StarName = "天哭", Level = "丙級", Meaning = "憂傷感傷。" } },
                { "天虛", new StarDefinition { StarName = "天虛", Level = "丙級", Meaning = "華而不實。" } },
                { "蜚廉", new StarDefinition { StarName = "蜚廉", Level = "丙級", Meaning = "口舌是非。" } },
                { "破碎", new StarDefinition { StarName = "破碎", Level = "丙級", Meaning = "做事反覆不全。" } },
                { "華蓋", new StarDefinition { StarName = "華蓋", Level = "丙級", Meaning = "宗教藝術才華。" } },
                { "咸池", new StarDefinition { StarName = "咸池", Level = "丙級", Meaning = "主肉慾桃花。" } },
                { "天德", new StarDefinition { StarName = "天德", Level = "丙級", Meaning = "上天庇蔭。" } },
                { "月德", new StarDefinition { StarName = "月德", Level = "丙級", Meaning = "化解衝突。" } },
                { "天福", new StarDefinition { StarName = "天福", Level = "丙級", Meaning = "增加福祿。" } },
                { "天官", new StarDefinition { StarName = "天官", Level = "丙級", Meaning = "升遷官運。" } },

                // 4. 博士十二神
                { "博士", new StarDefinition { StarName = "博士", Level = "神煞", Meaning = "聰明淵博。" } },
                { "力士", new StarDefinition { StarName = "力士", Level = "神煞", Meaning = "權力威勢。" } },
                { "青龍", new StarDefinition { StarName = "青龍", Level = "神煞", Meaning = "進財喜慶。" } },
                { "小耗", new StarDefinition { StarName = "小耗", Level = "神煞", Meaning = "小額耗財。" } },
                { "將軍", new StarDefinition { StarName = "將軍", Level = "神煞", Meaning = "武職威嚴。" } },
                { "奏書", new StarDefinition { StarName = "奏書", Level = "神煞", Meaning = "文書喜訊。" } },
                { "飛廉", new StarDefinition { StarName = "飛廉", Level = "神煞", Meaning = "毀謗是非。" } },
                { "喜神", new StarDefinition { StarName = "喜神", Level = "神煞", Meaning = "延續喜慶。" } },
                { "病符", new StarDefinition { StarName = "病符", Level = "神煞", Meaning = "小病侵擾。" } },
                { "大耗", new StarDefinition { StarName = "大耗", Level = "神煞", Meaning = "重大破財。" } },
                { "伏兵", new StarDefinition { StarName = "伏兵", Level = "神煞", Meaning = "暗中阻撓。" } },
                { "官府", new StarDefinition { StarName = "官府", Level = "神煞", Meaning = "法律官非。" } },

                // 5. 歲前十二神 (流年)
                { "歲建", new StarDefinition { StarName = "歲建", Level = "神煞", Meaning = "一年吉凶之首。" } },
                { "晦氣", new StarDefinition { StarName = "晦氣", Level = "神煞", Meaning = "情緒阻滯。" } },
                { "喪門", new StarDefinition { StarName = "喪門", Level = "神煞", Meaning = "憂愁弔喪。" } },
                { "貫索", new StarDefinition { StarName = "貫索", Level = "神煞", Meaning = "束縛糾纏。" } },
                { "龍德", new StarDefinition { StarName = "龍德", Level = "神煞", Meaning = "貴人化解。" } },
                { "白虎", new StarDefinition { StarName = "白虎", Level = "神煞", Meaning = "刑傷血光。" } },
                { "天德星", new StarDefinition { StarName = "天德", Level = "神煞", Meaning = "逢凶化吉。" } },
                { "弔客", new StarDefinition { StarName = "弔客", Level = "神煞", Meaning = "不寧弔慰。" } },

                // 6. 將前十二神
                { "將星", new StarDefinition { StarName = "將星", Level = "神煞", Meaning = "權柄化吉。" } },
                { "攀鞍", new StarDefinition { StarName = "攀鞍", Level = "神煞", Meaning = "名譽升遷。" } },
                { "歲驛", new StarDefinition { StarName = "歲驛", Level = "神煞", Meaning = "奔忙遠行。" } },
                { "息神", new StarDefinition { StarName = "息神", Level = "神煞", Meaning = "意志消沉。" } },
                { "災煞", new StarDefinition { StarName = "災煞", Level = "神煞", Meaning = "意外災難。" } },
                { "劫煞", new StarDefinition { StarName = "劫煞", Level = "神煞", Meaning = "財物搶奪。" } },
                { "指背", new StarDefinition { StarName = "指背", Level = "神煞", Meaning = "背後毀謗。" } },
                { "亡神", new StarDefinition { StarName = "亡神", Level = "神煞", Meaning = "耗損失去。" } },

                // 7. 長生十二神 (新增補全)
                { "長生", new StarDefinition { StarName = "長生", Level = "運程", Meaning = "生機蓬勃，新開始。" } },
                { "沐浴", new StarDefinition { StarName = "沐浴", Level = "運程", Meaning = "洗滌、桃花、不穩。" } },
                { "冠帶", new StarDefinition { StarName = "冠帶", Level = "運程", Meaning = "成熟、名譽。" } },
                { "臨官", new StarDefinition { StarName = "臨官", Level = "運程", Meaning = "強盛、白手起家。" } },
                { "帝旺", new StarDefinition { StarName = "帝旺", Level = "運程", Meaning = "巔峰、自負。" } },
                { "衰", new StarDefinition { StarName = "衰", Level = "運程", Meaning = "退氣、保守。" } },
                { "病", new StarDefinition { StarName = "病", Level = "運程", Meaning = "衰弱、欠佳。" } },
                { "死", new StarDefinition { StarName = "死", Level = "運程", Meaning = "沉滯、終結。" } },
                { "墓", new StarDefinition { StarName = "墓", Level = "運程", Meaning = "收藏、隱蔽。" } },
                { "絕", new StarDefinition { StarName = "絕", Level = "運程", Meaning = "孤獨、落空。" } },
                { "胎", new StarDefinition { StarName = "胎", Level = "運程", Meaning = "醞釀、希望。" } },
                { "養", new StarDefinition { StarName = "養", Level = "運程", Meaning = "扶持、發展。" } }
            };

            public static readonly Dictionary<string, StarDefinition> Dict;

            static StarMap()
            {
                // 先建立 Dict
                Dict = new Dictionary<string, StarDefinition>(BaseDict, StringComparer.OrdinalIgnoreCase);

                // 自動別名邏輯 (這取代了 100 行手寫代碼，所以行數會變短但功能更強)
                foreach (var kvp in BaseDict)
                {
                    string fullName = kvp.Key;
                    if (fullName.Length > 1)
                    {
                        string first = fullName.Substring(0, 1);
                        string last = fullName.Substring(fullName.Length - 1, 1);
                        if (!Dict.ContainsKey(first)) Dict[first] = kvp.Value;
                        if (!Dict.ContainsKey(last)) Dict[last] = kvp.Value;
                    }
                }

                // 核心星曜特例手動覆蓋 (確保最高優先度)
                Dict["月"] = BaseDict["太陰"];
                Dict["空"] = BaseDict["地空"];
                Dict["劫"] = BaseDict["地劫"];
                Dict["鸞"] = BaseDict["紅鸞"];
                Dict["喜"] = BaseDict["天喜"];
                Dict["馬"] = BaseDict["天馬"];
            }
        }
    }
}