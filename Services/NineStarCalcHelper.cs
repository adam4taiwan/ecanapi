namespace Ecanapi.Services
{
    /// <summary>九星氣學純計算層（不依賴 DB / Gemini，可被多個 Controller 共用）</summary>
    public static class NineStarCalcHelper
    {
        public static readonly string[] StarNames =
        {
            "", "一白水星", "二黑土星", "三碧木星", "四綠木星", "五黃土星",
            "六白金星", "七赤金星", "八白土星", "九紫火星"
        };

        // 星宿別名（貪狼/巨門/祿存/文曲/廉貞/武曲/破軍/左輔/右弼）
        public static readonly string[] StarAliases =
        {
            "", "貪狼星", "巨門星", "祿存星", "文曲星", "廉貞星",
            "武曲星", "破軍星", "左輔星", "右弼星"
        };

        public static readonly string[] StarDirections =
        {
            "", "北方", "西南方", "東方", "東南方", "中宮（以化解為主）", "西北方", "西方", "東北方", "南方"
        };

        public static readonly string[] StarColors =
        {
            "", "白色、藍色、灰色", "黃色、咖啡色", "青綠色", "翠綠色", "黃色",
            "白色", "白色、金色、銀色", "白色、淺黃色", "紅色、紫色"
        };

        public static readonly int[] StarLuckyNumbers = { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9 };

        // 星曜關鍵詞
        public static readonly string[] StarKeywords =
        {
            "", "研究、思考", "俸祿、成長", "明朗、前進", "信用、和諧", "權勢、統政",
            "活力、決斷", "魅力、交際", "儲蓄、改革", "名譽、改革"
        };

        // 得運描述
        public static readonly string[] StarProsper =
        {
            "",
            "得運時為官財星，主得名氣及官位，文武雙全，少年科甲官名，聲震四海。",
            "得運時位列尊崇，興隆置業，田產極盛，旺丁旺財，必生武貴，可出怪傑英豪。婦人當權，多謀節儉，可成霸業，亦可出醫師仙人。",
            "得運時興家立業，仕途官星大利，出法官律師及鬼才，特別大旺長房，必出刑貴或武貴。",
            "得運時為文昌星，大利文化藝術，科甲成名，進財進產，必得妻助或良夫，文章有價。",
            "得運時位處中極，威崇無比，恍若皇帝，攝盡四方，所以古代皇帝的龍袍都是黃色。",
            "得運時為財錦，此星一直認為是偏財橫財星，與一白、八白合成三大財星。生旺運時使人財兩旺，武貴義人出於此。",
            "得運時大利口才工作的人，旺歌星、演說家、占卜家等，得運時主旺財旺丁，大展拳腳，大力傳播通訊。",
            "得運時為太白財星，是一級財星，此星帶來功名富貴，置業成功，田宅科發，富不可當，為九星第一吉星。",
            "得運時為一級喜慶及愛情星，帶來良好姻緣和桃花，人緣良好，四海見利，旺丁旺財，必添良男賢女，亦為懷孕星，此星又興田產，大力置業及建築。"
        };

        // 失運描述
        public static readonly string[] StarDecline =
        {
            "",
            "失運時此星為桃花劫，因酒色財氣而破財損家，必患耳病或腎病，甚至性病，嚴重者必夫妻離異，孤身異鄉流亡。",
            "失運時為靈界，又名病符，一切最凶事均臨門生禍，死亡絕症，破產投死，寡婦當家，與五黃並列為最凶。",
            "失運時易招致刑險是非，小人當道，賊星入屋，破財招刑，官訴連年，此星易招膿血之病，足部手部頭髮肝膽絕症。此星好勇鬥狠，謠言誹謗之凶。",
            "失運時為桃花劫，必招酒色之禍，易招瘋哮血溢，肝膽或腰部以下出毛病，此星古代應驗於懸樑自盡，在現代社會應驗於服藥自殺，與藥品有關。",
            "失運時為五黃煞，為土煞之極，掌管死亡之事。嚴重時應驗死者為5人或5字之數，遇其它吉星可以化解一點凶性，但是如果遇上二黑之類，必患重病絕症。",
            "失運時必關乎血刃刀險，有刑險，對家宅影響必招致孤寡及血光意外，疾病驗在肺部。",
            "失運時主口舌是非，刀光劍影，開刀殘疾，凶在唇舌，橫死兵亂，世界大戰，牢獄刑險。又為火險之象，身體上影響呼吸、口舌和肺部。",
            "失運時為田產退讓，失財失義，小口損傷，瘟疫流行，手腳腰脊損傷，賭錢失利破家，破財於一瞬間。",
            "失運時為桃花劫星，主吐血及回祿之災，破財損丁於一瞬間，又主火災和爆炸，心臟及血崩等疾病，此星亦主眼病失明，火瘡流血。"
        };

        // 五行屬性
        private static readonly string[] StarElements =
        {
            "", "水", "土", "木", "木", "土", "金", "金", "土", "火"
        };

        // 甲子參考日
        private static readonly DateTime EpochDate = new DateTime(2000, 1, 1);
        private const int EpochCycleIndex = 10;

        /// <summary>計算本命星（考慮立春 2/4 前出生使用前一年）</summary>
        public static int CalcNatalStar(int year, int month, int day, int gender)
        {
            int y = (month < 2 || (month == 2 && day < 4)) ? year - 1 : year;
            return CalcYearStar(y, gender);
        }

        /// <summary>計算年飛星（中性，不分性別）</summary>
        public static int CalcYearStar(int year)
        {
            int last2 = year % 100;
            int sum = last2 / 10 + last2 % 10;
            while (sum >= 10) sum = sum / 10 + sum % 10;
            int star = year < 2000 ? (10 - sum) % 9 : (9 - sum) % 9;
            return star == 0 ? 9 : star;
        }

        /// <summary>計算年飛星（含性別）
        /// 男：1999前 = 10-S，2000後 = 9-S
        /// 女：1999前 = S-4（≤0則+9），2000後 = S-3（≤0則+9）
        /// S = 年尾兩位數字和（超過10再縮減至個位）
        /// </summary>
        public static int CalcYearStar(int year, int gender)
        {
            if (gender != 2) return CalcYearStar(year); // 男

            int last2 = year % 100;
            int sum = last2 / 10 + last2 % 10;
            while (sum >= 10) sum = sum / 10 + sum % 10;

            int star = year < 2000 ? (sum - 4) : (sum - 3);
            if (star <= 0) star += 9;
            return star;
        }

        /// <summary>計算月飛星</summary>
        public static int CalcMonthStar(DateTime date)
        {
            int yearStar = CalcYearStar(date.Year);
            int solarMonth = GetSolarMonth(date);
            int[] monthStartStars = { 0, 8, 5, 2, 8, 5, 2, 8, 5, 2 };
            int startStar = monthStartStars[yearStar];
            int star = ((startStar - solarMonth + 1) % 9 + 9) % 9;
            return star == 0 ? 9 : star;
        }

        /// <summary>計算日飛星</summary>
        public static int CalcDayStar(DateTime date)
        {
            var (startStar, forward, periodStart) = GetDayStarPeriod(date);
            DateTime jiazi = GetLastJiaZiDay(periodStart);
            int daysDiff = (int)(date.Date - jiazi).TotalDays;
            int star = forward
                ? ((startStar - 1 + daysDiff) % 9 + 9) % 9 + 1
                : ((startStar - 1 - daysDiff) % 9 + 9) % 9 + 1;
            return star;
        }

        /// <summary>計算時飛星（hour: 0-23）</summary>
        public static int CalcHourStar(DateTime dateTime)
        {
            int hour = dateTime.Hour;
            int branchIdx = hour == 23 ? 0 : hour / 2;
            int cycle = Get60CycleIndex(dateTime);
            int dayBranch = cycle % 12;
            int[] meng = { 2, 5, 8, 11 };
            int[] zhong = { 0, 3, 6, 9 };
            bool isYangHalf = IsYangHalf(dateTime);
            int startStar;
            if (meng.Contains(dayBranch))
                startStar = isYangHalf ? 7 : 3;
            else if (zhong.Contains(dayBranch))
                startStar = isYangHalf ? 1 : 9;
            else
                startStar = isYangHalf ? 4 : 6;
            int star = ((startStar - 1 + branchIdx) % 9 + 9) % 9 + 1;
            return star;
        }

        public static string GetStarElement(int star) =>
            (star >= 1 && star <= 9) ? StarElements[star] : "";

        /// <summary>
        /// 計算主位A × 勳位B 的五行基礎關係與吉凶
        /// 返回 (relation, verdict, note)
        /// relation: 印生/官克/泄氣/得財/比旺
        /// verdict:  大吉/吉/偏吉/偏凶/凶/特殊（同星）
        /// </summary>
        public static (string relation, string verdict, string note) CalcFiveElementCombination(int starA, int starB)
        {
            if (starA < 1 || starA > 9 || starB < 1 || starB > 9)
                return ("", "平", "");

            if (starA == starB)
                return ("比旺", "特殊", $"同星相遇（{StarNames[starA]}），得運則最吉，失運則最凶");

            string eA = StarElements[starA];
            string eB = StarElements[starB];

            // B生A（勳位生主位）= 印生，吉
            if (Generates(eB, eA))
                return ("印生", "吉", $"{StarNames[starB]}（{eB}）生{StarNames[starA]}（{eA}），主位得助");

            // B克A（勳位克主位）= 官克，凶
            if (Controls(eB, eA))
                return ("官克", "凶", $"{StarNames[starB]}（{eB}）克{StarNames[starA]}（{eA}），主位受制");

            // A生B（主位生勳位）= 泄氣，偏凶
            if (Generates(eA, eB))
                return ("泄氣", "偏凶", $"{StarNames[starA]}（{eA}）生{StarNames[starB]}（{eB}），主位泄氣耗損");

            // A克B（主位克勳位）= 得財，偏吉
            if (Controls(eA, eB))
                return ("得財", "偏吉", $"{StarNames[starA]}（{eA}）克{StarNames[starB]}（{eB}），主位得財");

            return ("平", "平", "");
        }

        /// <summary>依得運/失運狀態修正五行組合吉凶</summary>
        public static string ApplyYunModifier(string verdict, bool isProspering)
        {
            return (verdict, isProspering) switch
            {
                ("大吉", true)  => "大吉",
                ("吉",   true)  => "大吉",
                ("偏吉", true)  => "吉",
                ("平",   true)  => "偏吉",
                ("偏凶", true)  => "平（凶減半）",
                ("凶",   true)  => "偏凶（凶減半）",
                ("大凶", true)  => "凶（凶減半）",
                ("大吉", false) => "吉（失運打折）",
                ("吉",   false) => "偏吉（失運打折）",
                ("偏吉", false) => "平（失運打折）",
                ("平",   false) => "偏凶",
                ("偏凶", false) => "凶",
                ("凶",   false) => "大凶",
                ("大凶", false) => "大凶",
                _              => verdict
            };
        }

        // ── 五行相生相克 ──
        private static bool Generates(string from, string to) =>
            (from == "水" && to == "木") || (from == "木" && to == "火") ||
            (from == "火" && to == "土") || (from == "土" && to == "金") ||
            (from == "金" && to == "水");

        private static bool Controls(string from, string to) =>
            (from == "水" && to == "火") || (from == "火" && to == "金") ||
            (from == "金" && to == "木") || (from == "木" && to == "土") ||
            (from == "土" && to == "水");

        /// <summary>依年份取得當前三元九運（每運20年）</summary>
        public static int GetCurrentYun(int year)
        {
            if (year >= 2044) return 1;
            if (year >= 2024) return 9;
            if (year >= 2004) return 8;
            if (year >= 1984) return 7;
            if (year >= 1964) return 6;
            if (year >= 1944) return 5;
            if (year >= 1924) return 4;
            if (year >= 1904) return 3;
            if (year >= 1884) return 2;
            return 1;
        }

        /// <summary>
        /// 判斷本命星在當前運的狀態：
        /// 旺=當運，生=下運，較遠生氣=再下運，方衰=上運，更衰=再上運，死氣=其餘
        /// 返回 (狀態標籤, 是否得運)
        /// </summary>
        public static (string label, bool isProspering) GetStarYunStatus(int natalStar, int currentYun)
        {
            int next1 = currentYun % 9 + 1;
            int next2 = next1 % 9 + 1;
            int prev1 = currentYun == 1 ? 9 : currentYun - 1;
            int prev2 = prev1 == 1 ? 9 : prev1 - 1;

            if (natalStar == currentYun) return ("旺（當運最旺，得運）", true);
            if (natalStar == next1)      return ("生（下運生氣，得運）", true);
            if (natalStar == next2)      return ("較遠生氣（可小用）", true);
            if (natalStar == prev1)      return ("方衰（剛過，尚可）", false);
            if (natalStar == prev2)      return ("更衰（失運）", false);
            return ("死氣（失運最重）", false);
        }

        /// <summary>
        /// 流年飛宮：通用年星飛入中宮（5宮）後，計算飛入指定宮位（本命星=宮位號）的星
        /// 公式：((universalYearStar + palace - 6) % 9 + 9) % 9 + 1
        /// </summary>
        public static int CalcFlyingStarInPalace(int universalYearStar, int palace)
            => ((universalYearStar + palace - 6) % 9 + 9) % 9 + 1;

        // ── 內部輔助 ──

        private static int Get60CycleIndex(DateTime date)
        {
            int days = (int)(date.Date - EpochDate).TotalDays;
            return ((EpochCycleIndex + days) % 60 + 60) % 60;
        }

        private static DateTime GetLastJiaZiDay(DateTime from)
        {
            int idx = Get60CycleIndex(from);
            return from.Date.AddDays(-idx);
        }

        private static (int startStar, bool forward, DateTime periodStart) GetDayStarPeriod(DateTime date)
        {
            int year = date.Year;
            int doy = date.DayOfYear;
            int doyYuShui = 50;
            int doyGuYu = 110;
            int doyXiaZhi = 172;
            int doyChuShu = 235;
            int doyShuangJiang = 296;
            int doyDongZhi = DateTime.IsLeapYear(year) ? 357 : 356;

            if (doy < doyYuShui) return (1, true, new DateTime(year - 1, 12, 22));
            if (doy < doyGuYu) return (7, true, new DateTime(year, 2, 19));
            if (doy < doyXiaZhi) return (4, true, new DateTime(year, 4, 20));
            if (doy < doyChuShu) return (9, false, new DateTime(year, 6, 21));
            if (doy < doyShuangJiang) return (3, false, new DateTime(year, 8, 23));
            if (doy < doyDongZhi) return (6, false, new DateTime(year, 10, 23));
            return (1, true, new DateTime(year, 12, 22));
        }

        private static bool IsYangHalf(DateTime date)
        {
            int doy = date.DayOfYear;
            int doyXiaZhi = DateTime.IsLeapYear(date.Year) ? 173 : 172;
            int doyDongZhi = DateTime.IsLeapYear(date.Year) ? 357 : 356;
            return doy >= doyDongZhi || doy < doyXiaZhi;
        }

        private static int GetSolarMonth(DateTime date)
        {
            int m = date.Month, d = date.Day;
            (int sm, int sd, int solarM)[] terms =
            {
                (1, 6, 12), (2, 4, 1), (3, 6, 2), (4, 5, 3),
                (5, 6, 4), (6, 6, 5), (7, 7, 6), (8, 7, 7),
                (9, 8, 8), (10, 8, 9), (11, 7, 10), (12, 7, 11),
            };
            int current = 11;
            foreach (var (sm, sd, solarM) in terms)
                if (m > sm || (m == sm && d >= sd)) current = solarM;
            return current;
        }
    }
}
