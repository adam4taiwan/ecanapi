namespace Ecanapi.Services.AstrologyEngine
{
    /// <summary>
    /// 提供命理計算中常用的小工具函式
    /// </summary>
    public static class AstrologyHelper
    {
        /// <summary>
        /// 將天干文字轉換為數字 (甲=1, 乙=2, ...)
        /// </summary>
        public static int SkyToNumber(string sky)
        {
            return Array.IndexOf(AstrologyConstants.S_SKY, sky);
        }

        /// <summary>
        /// 將地支文字轉換為數字 (子=1, 丑=2, ...)
        /// </summary>
        public static int FloorToNumber(string floor)
        {
            return Array.IndexOf(AstrologyConstants.S_FLO, floor);
        }

        /// <summary>
        /// 將時辰干支（例如："甲子"）轉換為地支對應的數字 (子=1, 丑=2, ...)
        /// </summary>
        public static int GetChineseHourValue(string chineseHour)
        {
            if (string.IsNullOrEmpty(chineseHour) || chineseHour.Length < 2)
            {
                return 1; // 預設為子時
            }
            char branch = chineseHour[1]; // 取地支的字
            switch (branch)
            {
                case '子': return 1;
                case '丑': return 2;
                case '寅': return 3;
                case '卯': return 4;
                case '辰': return 5;
                case '巳': return 6;
                case '午': return 7;
                case '未': return 8;
                case '申': return 9;
                case '酉': return 10;
                case '戌': return 11;
                case '亥': return 12;
                default: return 1;
            }
        }

        // 檔案: AstrologyHelper.cs

        // ... (省略現有的 AstrologyHelper 類別)

        // 【新增靜態類別】: 用於計算固定位置的貴人星和祿羊陀
        public static class MinorStarPlacer
        {
            /// <summary>
            /// 根據年干計算固定位置的星曜 (天魁, 天鉞, 祿存, 擎羊, 陀羅)。
            /// </summary>
            /// <param name="yearStem">年干文字 (e.g., "甲")</param>
            /// <returns>Dictionary<int, string>，Key=地支索引(1=子...12=亥)，Value=星曜名稱(例如 "魁" 或 "祿|羊")</returns>
            public static Dictionary<int, string> CalculateFixedNobleAndLuYangTuo(string yearStem)
            {
                var placements = new Dictionary<int, string>();
                int yearStemNum = AstrologyHelper.SkyToNumber(yearStem);

                // --- 1. 天魁 (Kwai) / 天鉞 (Yue) - 年干起貴人 ---
                // 規則來自貴人星及雜星規則.txt 的前段
                int kwaiPos = 0; // 天魁 (貴人)
                int yuePos = 0;  // 天鉞 (天乙貴人)

                if (yearStemNum == 1 || yearStemNum == 5 || yearStemNum == 7) // 甲(1), 戊(5), 庚(7) -> 丑(2), 未(8)
                { kwaiPos = 2; yuePos = 8; }
                else if (yearStemNum == 2 || yearStemNum == 6) // 乙(2), 己(6) -> 子(1), 申(9)
                { kwaiPos = 1; yuePos = 9; }
                else if (yearStemNum == 3 || yearStemNum == 4) // 丙(3), 丁(4) -> 亥(12), 酉(10)
                { kwaiPos = 12; yuePos = 10; }
                else if (yearStemNum == 9 || yearStemNum == 10) // 壬(9), 癸(10) -> 卯(4), 巳(6)
                { kwaiPos = 4; yuePos = 6; }
                else if (yearStemNum == 8) // 辛(8) -> 午(7), 寅(3)
                { kwaiPos = 7; yuePos = 3; }

                // 確保不覆蓋
                placements[kwaiPos] = placements.ContainsKey(kwaiPos) ? placements[kwaiPos] + "|魁" : "魁";
                placements[yuePos] = placements.ContainsKey(yuePos) ? placements[yuePos] + "|鉞" : "鉞";

                // --- 2. 祿存 (Lu), 擎羊 (Yang), 陀羅 (Tuo) - 年干起四煞 ---
                // 規則來自貴人星及雜星規則.txt
                int luPos = 0; // 祿存

                switch (yearStemNum)
                {
                    case 1: luPos = 3; break; // 甲祿在寅(3)
                    case 2: luPos = 4; break; // 乙祿在卯(4)
                    case 3: luPos = 6; break; // 丙祿在巳(6)
                    case 4: luPos = 7; break; // 丁祿在午(7)
                    case 5: luPos = 6; break; // 戊祿在巳(6)
                    case 6: luPos = 7; break; // 己祿在午(7)
                    case 7: luPos = 9; break; // 庚祿在申(9)
                    case 8: luPos = 10; break; // 辛祿在酉(10)
                    case 9: luPos = 12; break; // 壬祿在亥(12)
                    case 10: luPos = 1; break; // 癸祿在子(1)
                }

                // 擎羊 (Yang) 在 祿存的下一宮 (順時針)
                int yangPos = luPos % 12 + 1;

                // 陀羅 (Tuo) 在 祿存的上一宮 (逆時針)
                int tuoPos = luPos - 1;
                if (tuoPos == 0) tuoPos = 12;

                placements[luPos] = placements.ContainsKey(luPos) ? placements[luPos] + "|祿" : "祿";
                placements[yangPos] = placements.ContainsKey(yangPos) ? placements[yangPos] + "|羊" : "羊";
                placements[tuoPos] = placements.ContainsKey(tuoPos) ? placements[tuoPos] + "|陀" : "陀";

                return placements; // Key: 地支索引 (1=子, ..., 12=亥), Value: 星曜名稱
            }

            // *** 由於火星/鈴星的計算邏輯過於複雜，且涉及到您未提供的 CRC 數組，我們將暫時依賴舊引擎的結果，
            //     只修正因欄位錯位導致的錯誤。如果問題仍存在，則需要提供 CRC 的計算邏輯。***
        }
    }
}