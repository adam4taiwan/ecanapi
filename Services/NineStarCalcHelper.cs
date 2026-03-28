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

        public static readonly string[] StarDirections =
        {
            "", "北方", "西南方", "東方", "東南方", "中宮（以化解為主）", "西北方", "西方", "東北方", "南方"
        };

        public static readonly string[] StarColors =
        {
            "", "白色、藍色", "黃色、棕色", "綠色、青色", "綠色、青色", "黃色（需化解）",
            "白色、金色", "白色、金色", "白色、黃色", "紫色、紅色"
        };

        public static readonly int[] StarLuckyNumbers = { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9 };

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
