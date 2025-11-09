using System;

namespace Print2Engine
{
    public class EcanCalendarAdapter : IEcanCalendar
    {
        public string GetChineseDate(DateTime dt)
        {
            var cc = new Ecan.EcanChineseCalendar(dt.Year, dt.Month, dt.Day, false);
            return cc.ChineseDateString;
        }

        public string GetChineseZodiac(DateTime dt)
        {
            var cc = new Ecan.EcanChineseCalendar(dt.Year, dt.Month, dt.Day, false);
            return cc.AnimalString;
        }

        public string GetSolarTerm(DateTime dt)
        {
            return GetNearestSolarTerm(dt);
        }

        public string GetPreviousSolarTerm(DateTime dt)
        {
            return GetNearestSolarTerm(dt, previous: true);
        }

        public string GetNextSolarTerm(DateTime dt)
        {
            return GetNearestSolarTerm(dt, previous: false);
        }

        public string GetGanZhi(DateTime dt)
        {
            var cc = new Ecan.EcanChineseCalendar(dt.Year, dt.Month, dt.Day, false);

            string yearGanZhi = cc.GanZhiDateString.Substring(0, 2);  // 年柱
            string monthGanZhi = cc.GanZhiDateString.Substring(3, 2); // 月柱
            string dayGanZhi = cc.GanZhiDateString.Substring(6, 2);   // 日柱
            string hourGanZhi = GetHourGanZhi(cc, dt.Hour);           // 時柱
            //hourGanZhi = cc.ChineseHour;           // 時柱
            //return $"{yearGanZhi}{monthGanZhi}{dayGanZhi}";
            return $"{yearGanZhi}{monthGanZhi}{dayGanZhi}{hourGanZhi}";
        }

        private string GetHourGanZhi(Ecan.EcanChineseCalendar cc, int hour)
        {
            // 地支
            string[] dizhi = { "子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥" };
            string[] tiangan = { "甲", "乙", "丙", "丁", "戊", "己", "庚", "辛", "壬", "癸" };

            int index = (hour + 1) / 2 % 12;
            string dz = dizhi[index];

            string dayGan = cc.GanZhiDateString.Substring(6, 1);
            int dayGanIndex = Array.IndexOf(tiangan, dayGan);
            if (dayGanIndex < 0) dayGanIndex = 0;

            int tgIndex = (dayGanIndex * 2 + index) % 10;
            string tg = tiangan[tgIndex];

            return tg + dz;
        }

        //private string GetNearestSolarTerm(DateTime dt, bool? previous = null)
        //{
        //    // 這裡我們只能用 EcanChineseCalendar 的 SolarTerm 計算
        //    var cc = new Ecan.EcanChineseCalendar(dt.Year, dt.Month, dt.Day, false);

        //    // ⚠️ 假設 EcanChineseCalendar 提供了計算節氣日期的方法，否則需要自己加演算法
        //    // 為了範例，先回傳空字串，避免編譯錯誤
        //    return string.Empty;
        //}
        private string GetNearestSolarTerm(DateTime dt, bool? previous = null)
        {
            var cc = new Ecan.EcanChineseCalendar(dt.Year, dt.Month, dt.Day, false);

            if (previous.HasValue)
            {
                if (previous.Value)
                {
                    // 取前一個節氣
                    return cc.ChineseTwentyFourPrevDay().ToString("yyyy-MM-dd");
                }
                else
                {
                    // 取下一個節氣
                    return cc.ChineseTwentyFourNext().ToString("yyyy-MM-dd");
                }
            }
            else
            {
                // 取最接近的節氣
                var prevTerm = cc.ChineseTwentyFourPrevDay();
                var nextTerm = cc.ChineseTwentyFourNext();

                var diffPrev = Math.Abs((dt - prevTerm).TotalDays);
                var diffNext = Math.Abs((dt - nextTerm).TotalDays);

                if (diffPrev <= diffNext)
                {
                    return prevTerm.ToString("yyyy-MM-dd");
                }
                else
                {
                    return nextTerm.ToString("yyyy-MM-dd");
                }
            }
        }
        public string GetConstellation(DateTime dt)
        {
            var cc = new Ecan.EcanChineseCalendar(dt.Year, dt.Month, dt.Day, false);
            return cc.Constellation;
        }

        public string GetChineseConstellation(DateTime dt)
        {
            var cc = new Ecan.EcanChineseCalendar(dt.Year, dt.Month, dt.Day, false);
            return cc.ChineseConstellation;
        }
    }
}
