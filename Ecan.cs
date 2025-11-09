
using System;
using System.Collections.Generic;
using System.Text;

namespace Ecan
{
    #region ChineseCalendarException
    /// <summary>
    /// 中?日?異常?理
    /// </summary>
    public class newCalendarException : System.Exception
    {
        public newCalendarException(string msg)
            : base(msg)
        {
        }
    }

    #endregion

    public class EcanChineseCalendar
    {
         #region ?部?構
        private struct SolarHolidayStruct//??
        {
            public int Month;
            public int Day;
            public int Recess; //假期?度
            public string HolidayName;
            public SolarHolidayStruct(int month, int day, int recess, string name)
            {
                Month = month;
                Day = day;
                Recess = recess;
                HolidayName = name;
            }
        }

        private struct LunarHolidayStruct//??
        {
            public int Month;
            public int Day;
            public int Recess;
            public string HolidayName;

            public LunarHolidayStruct(int month, int day, int recess, string name)
            {
                Month = month;
                Day = day;
                Recess = recess;
                HolidayName = name;
            }
        }

        private struct WeekHolidayStruct
        {
            public int Month;
            public int WeekAtMonth;
            public int WeekDay;
            public string HolidayName;

            public WeekHolidayStruct(int month, int weekAtMonth, int weekDay, string name)
            {
                Month = month;
                WeekAtMonth = weekAtMonth;
                WeekDay = weekDay;
                HolidayName = name;
            }
        }
        #endregion

        #region ?部?量
        private DateTime _date;
        private DateTime _datetime;

        private int _cYear;
        private int _cMonth;
        private int _cDay;
        private bool _cIsLeapMonth; //?月是否?月
        private bool _cIsLeapYear; //?年是否有?月
        #endregion

        #region 基??據
        #region 基本常量
        private const int MinYear = 1900;
        private const int MaxYear = 2050;
        private static DateTime MinDay = new DateTime(1900, 1, 30);
        private static DateTime MaxDay = new DateTime(2049, 12, 31);
        public const int GanZhiStartYear = 1864; //干支?算起始年
        public static DateTime GanZhiStartDay = new DateTime(1899, 12, 22);//起始日
        private const string HZNum = "零一二三四五六七八九";
        private const int AnimalStartYear = 1900; //1900年?鼠年
        public static DateTime ChineseConstellationReferDay = new DateTime(2007, 9, 13);//28星宿?考值,本日?角
        #endregion

        #region ???據
        /// <summary>
        /// ?源於網上的???據
        /// </summary>
        /// <remarks>
        /// ?據?構如下，共使用17位?據
        /// 第17位：表示?月天?，0表示29天   1表示30天
        /// 第16位-第5位（共12位）表示12?月，其中第16位表示第一月，如果?月?30天??1，29天?0
        /// 第4位-第1位（共4位）表示?月是哪?月，如果?年?有?月，?置0
        ///</remarks>
        private static int[] LunarDateArray = new int[]{
                0x04BD8,0x04AE0,0x0A570,0x054D5,0x0D260,0x0D950,0x16554,0x056A0,0x09AD0,0x055D2,
                0x04AE0,0x0A5B6,0x0A4D0,0x0D250,0x1D255,0x0B540,0x0D6A0,0x0ADA2,0x095B0,0x14977,
                0x04970,0x0A4B0,0x0B4B5,0x06A50,0x06D40,0x1AB54,0x02B60,0x09570,0x052F2,0x04970,
                0x06566,0x0D4A0,0x0EA50,0x06E95,0x05AD0,0x02B60,0x186E3,0x092E0,0x1C8D7,0x0C950,
                0x0D4A0,0x1D8A6,0x0B550,0x056A0,0x1A5B4,0x025D0,0x092D0,0x0D2B2,0x0A950,0x0B557,
                0x06CA0,0x0B550,0x15355,0x04DA0,0x0A5B0,0x14573,0x052B0,0x0A9A8,0x0E950,0x06AA0,
                0x0AEA6,0x0AB50,0x04B60,0x0AAE4,0x0A570,0x05260,0x0F263,0x0D950,0x05B57,0x056A0,
                0x096D0,0x04DD5,0x04AD0,0x0A4D0,0x0D4D4,0x0D250,0x0D558,0x0B540,0x0B6A0,0x195A6,
                0x095B0,0x049B0,0x0A974,0x0A4B0,0x0B27A,0x06A50,0x06D40,0x0AF46,0x0AB60,0x09570,
                0x04AF5,0x04970,0x064B0,0x074A3,0x0EA50,0x06B58,0x055C0,0x0AB60,0x096D5,0x092E0,
                0x0C960,0x0D954,0x0D4A0,0x0DA50,0x07552,0x056A0,0x0ABB7,0x025D0,0x092D0,0x0CAB5,
                0x0A950,0x0B4A0,0x0BAA4,0x0AD50,0x055D9,0x04BA0,0x0A5B0,0x15176,0x052B0,0x0A930,
                0x07954,0x06AA0,0x0AD50,0x05B52,0x04B60,0x0A6E6,0x0A4E0,0x0D260,0x0EA65,0x0D530,
                0x05AA0,0x076A3,0x096D0,0x04BD7,0x04AD0,0x0A4D0,0x1D0B6,0x0D250,0x0D520,0x0DD45,
                0x0B5A0,0x056D0,0x055B2,0x049B0,0x0A577,0x0A4B0,0x0AA50,0x1B255,0x06D20,0x0ADA0,
                0x14B63        
                };

        #endregion

        #region 星座名?
        private static string[] _constellationName = 
                { 
                    "白羊座", "金牛座", "雙子座", 
                    "巨蟹座", "獅子座", "處女座", 
                    "天秤座", "天蠍座", "射手座", 
                    "摩羯座", "水瓶座", "雙魚座"
                };
        #endregion

        #region 二十四?氣
        private static string[] _lunarHolidayName = 
                    { 
                    "小寒", "大寒", "立春", "雨水", 
                    "驚蟄", "春分", "清明", "穀雨", 
                    "立夏", "小滿", "芒種", "夏至", 
                    "小暑", "大暑", "立秋", "處暑", 
                    "白露", "秋分", "寒露", "霜降", 
                    "立冬", "小雪", "大雪", "冬至"
                    };
        #endregion

        #region 二十八星宿
        public static string[] _chineseConstellationName =
            {
                  //四        五      六         日        一      二      三  
                "角木蛟","亢金龍","女土蝠","房日兔","心月狐","尾火虎","箕水豹",
                "鬥木獬","牛金牛","氐土貉","虛日鼠","危月燕","室火豬","壁水貐",
                "奎木狼","婁金狗","胃土雉","昴日雞","畢月烏","觜火猴","參水猿",
                "井木犴","鬼金羊","柳土獐","星日馬","張月鹿","翼火蛇","軫水蚓" 
            };
        #endregion

        #region ?氣?據
        public static string[] SolarTerm = new string[] { "小寒", "大寒", "立春", "雨水", "驚蟄", "春分", "清明", "穀雨", "立夏", "小滿", "芒種", "夏至", "小暑", "大暑", "立秋", "處暑", "白露", "秋分", "寒露", "霜降", "立冬", "小雪", "大雪", "冬至" };
        //public static string[] SolarTerm = new string[] { "大寒", "立春", "雨水", "驚蟄", "春分", "清明", "穀雨", "立夏", "小滿", "芒種", "夏至", "小暑", "大暑", "立秋", "處暑", "白露", "秋分", "寒露", "霜降", "立冬", "小雪", "大雪", "冬至" ,"小寒", "大寒" };
        private static int[] sTermInfo = new int[] { 0, 21208, 42467, 63836, 85337, 107014, 128867, 150921, 173149, 195551, 218072, 240693, 263343, 285989, 308563, 331033, 353350, 375494, 397447, 419210, 440795, 462224, 483532, 504758 };
        #endregion

        #region ??相??據
        private static string ganStr = "甲乙丙丁戊己庚辛壬癸";
        private static string zhiStr = "子丑寅卯辰巳午未申酉戌亥";
        private static string animalStr = "鼠牛虎兔龍蛇馬羊猴雞狗豬";
        private static string nStr1 = "日一二三四五六七八九";
        private static string nStr2 = "初十廿卅";
        private static string[] _monthString =
                {
                    "出?","正月","二月","三月","四月","五月","六月","七月","八月","九月","十月","十一月","十二月"
                };
        #endregion
        /// <summary>
        /// 獲取節氣
        /// </summary>
        public string SolarTermString
        {
            get
            {
                if (_cYear == 0)
                    ComputeChineseDate();
                return getSolarTerm(_date);
            }
        }
        private static DateTime SolarTermDate(int year, int index)
        {
            int y = year - 1900;
            double d = (y * 365.2422 + sTermInfo[index] / 1000);
            DateTime dt = new DateTime(year, 1, 1).AddDays(d);
            return dt;
        }
        /// <summary>
        /// 計算節氣
        /// </summary>
        /// <param name="date">日期</param>
        /// <returns></returns>
        private static string getSolarTerm(DateTime date)
        {
            if (date.Year < MinYear || date.Year > MaxYear)
                return string.Empty;
            string str = string.Empty;
            int index = -1;
            for (int i = 0; i < 24; i++)
            {
                if (SolarTermDate(date.Year, i).Day == date.Day)
                {
                    index = i;
                    break;
                }
            }
            if (index != -1)
                str = SolarTerm[index];
            return str;
        }
        /// <summary>
        /// 計算?年是否?有?月
        /// </summary>
        /// <param name="year">年份</param>
        /// <returns></returns>
        private static int leapMonth(int year)
        {
            return LunarDateArray[year - MinYear] & 0xf;
        }
        /// 計算?年?月的天?
        /// </summary>
        /// <param name="year">年份</param>
        /// <param name="isLeap">是否?月</param>
        /// <returns></returns>
        private static int leapMonthDays(int year)
        {
            int month = leapMonth(year);
            if (month == 0)
                return 0;
            return (int)((LunarDateArray[year - MinYear] & (0x80000 >> month)) != 0 ? 30 : 29);
        }
        /// <summary>
        /// 計算?年共有多少天
        /// </summary>
        /// <param name="year">年份</param>
        /// <returns></returns>
        private static int yearDays(int year)
        {
            int i, sum = 348;
            for (i = 0x80000; i > 0x8; i >>= 1)
            {
                if ((LunarDateArray[year - MinYear] & i) != 0)
                    sum++;
            }
            return sum + leapMonthDays(year);
        }
        private void ComputeChineseDate()
        {
            _cIsLeapYear = false;
            int offset = _date.Subtract(MinDay).Days;
            int i, sum = 0;
            for (i = MinYear; i <= MaxYear; i++)
            {
                sum = yearDays(i);
                if (offset < sum)
                    break;
                offset -= sum;
            }
            _cYear = i;
            int leap = leapMonth(_cYear); //?年?月
            _cIsLeapYear = leap > 0 ? true : false;
            for (i = 1; i <= 12; i++)
            {
                if (leap > 0 && i == leap + 1 && !_cIsLeapMonth)
                {
                    if (offset < leapMonthDays(_cYear))
                    {
                        _cIsLeapMonth = true;
                        break;
                    }
                    offset -= leapMonthDays(_cYear);
                }
                offset -= monthDays(_cYear, i);
                if (offset < 0)
                    break;
            }
            _cMonth = i;
            _cDay = offset + 1;
        }

        /// <summary>
        /// 計算?年?月的天?
        /// </summary>
        /// <param name="year">年份</param>
        /// <param name="month">月份</param>
        /// <returns></returns>
        private static int monthDays(int year, int month)
        {
            if (month < 1 || month > 12)
                return 0;
            return (int)((LunarDateArray[year - MinYear] & (0x80000 >> (month - 1))) != 0 ? 30 : 29);
        }
        public string GetRealMonth(DateTime date1)
        {
            //*********************************************************************************
            // 節氣無任何確定規律,所以只好建立表格對應
            //**********************************************************************************}
            // 數據格式說明:
            // 如 1901 年的節氣為
            // 1月 2月 3月 4月 5月 6月 7月 8月 9月 10月 11月 12月
            // 6, 21, 4, 19, 6, 21, 5, 21, 6,22, 6,22, 8, 23, 8, 24, 8, 24, 8, 24, 8, 23, 8, 22
            // 9, 6, 11,4, 9, 6, 10,6, 9,7, 9,7, 7, 8, 7, 9, 7, 9, 7, 9, 7, 8, 7, 15
            // 上面第一行數據為每月節氣對應日期,15減去每月第一個節氣,每月第二個節氣減去15得第二行
            // 這樣每月兩個節氣對應數據都小於16,每月用一個陣列元素存放
            // 高位元(high-bit)存放第一個節氣數據,低位元(low-bit)存放第二個節氣的數據,可得下表
            byte[] gLunarHoliDay ={
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x79, 0x69, 0x78, 0x77, //1901
              0x96, 0xA4, 0x96, 0x96, 0x97, 0x87, 0x79, 0x79, 0x79, 0x69, 0x78, 0x78, //1902
              0x96, 0xA5, 0x87, 0x96, 0x87, 0x87, 0x79, 0x69, 0x69, 0x69, 0x78, 0x78, //1903
              0x86, 0xA5, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x79, 0x78, 0x87, //1904
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x79, 0x69, 0x78, 0x77, //1905
              0x96, 0xA4, 0x96, 0x96, 0x97, 0x97, 0x79, 0x79, 0x79, 0x69, 0x78, 0x78, //1906
              0x96, 0xA5, 0x87, 0x96, 0x87, 0x87, 0x79, 0x69, 0x69, 0x69, 0x78, 0x78, //1907
              0x86, 0xA5, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x69, 0x78, 0x87, //1908
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x79, 0x69, 0x78, 0x77, //1909
              0x96, 0xA4, 0x96, 0x96, 0x97, 0x97, 0x79, 0x79, 0x79, 0x69, 0x78, 0x78, //1910
              0x96, 0xA5, 0x87, 0x96, 0x87, 0x87, 0x79, 0x69, 0x69, 0x69, 0x78, 0x78, //1911
              0x86, 0xA5, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x69, 0x78, 0x87, //1912
              0x95, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x79, 0x69, 0x78, 0x77, //1913
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x79, 0x79, 0x79, 0x69, 0x78, 0x78, //1914
              0x96, 0xA5, 0x97, 0x96, 0x97, 0x87, 0x79, 0x79, 0x69, 0x69, 0x78, 0x78, //1915
              0x96, 0xA5, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x79, 0x77, 0x87, //1916
              0x95, 0xB4, 0x96, 0xA6, 0x96, 0x97, 0x78, 0x79, 0x78, 0x69, 0x78, 0x87, //1917
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x79, 0x79, 0x79, 0x69, 0x78, 0x77, //1918
              0x96, 0xA5, 0x97, 0x96, 0x97, 0x87, 0x79, 0x79, 0x69, 0x69, 0x78, 0x78, //1919
              0x96, 0xA5, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x79, 0x77, 0x87, //1920
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x78, 0x79, 0x78, 0x69, 0x78, 0x87, //1921
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x79, 0x79, 0x79, 0x69, 0x78, 0x77, //1922
              0x96, 0xA4, 0x96, 0x96, 0x97, 0x87, 0x79, 0x79, 0x69, 0x69, 0x78, 0x78, //1923
              0x96, 0xA5, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x79, 0x77, 0x87, //1924
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x78, 0x79, 0x78, 0x69, 0x78, 0x87, //1925
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x79, 0x69, 0x78, 0x77, //1926
              0x96, 0xA4, 0x96, 0x96, 0x97, 0x87, 0x79, 0x79, 0x79, 0x69, 0x78, 0x78, //1927
              0x96, 0xA5, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x78, 0x87, 0x87, //1928
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x79, 0x77, 0x87, //1929
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x79, 0x69, 0x78, 0x77, //1930
              0x96, 0xA4, 0x96, 0x96, 0x97, 0x87, 0x79, 0x79, 0x79, 0x69, 0x78, 0x78, //1931
              0x96, 0xA5, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x78, 0x87, 0x87, //1932
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x69, 0x78, 0x87, //1933
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x79, 0x69, 0x78, 0x77, //1934
              0x96, 0xA4, 0x96, 0x96, 0x97, 0x97, 0x79, 0x79, 0x79, 0x69, 0x78, 0x78, //1935
              0x96, 0xA5, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x78, 0x87, 0x87, //1936
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x69, 0x78, 0x87, //1937
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x79, 0x69, 0x78, 0x77, //1938
              0x96, 0xA4, 0x96, 0x96, 0x97, 0x97, 0x79, 0x79, 0x79, 0x69, 0x78, 0x78, //1939
              0x96, 0xA5, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x78, 0x87, 0x87, //1940
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x69, 0x78, 0x87, //1941
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x79, 0x69, 0x78, 0x77, //1942
              0x96, 0xA4, 0x96, 0x96, 0x97, 0x97, 0x79, 0x79, 0x79, 0x69, 0x78, 0x78, //1943
              0x96, 0xA5, 0x96, 0xA5, 0xA6, 0x96, 0x88, 0x78, 0x78, 0x78, 0x87, 0x87, //1944
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x79, 0x77, 0x87, //1945
              0x95, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x78, 0x69, 0x78, 0x77, //1946
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x79, 0x79, 0x79, 0x69, 0x78, 0x78, //1947
              0x96, 0xA5, 0xA6, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x78, 0x78, 0x87, 0x87, //1948
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x79, 0x78, 0x79, 0x77, 0x87, //1949
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x78, 0x79, 0x78, 0x69, 0x78, 0x77, //1950
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x79, 0x79, 0x79, 0x69, 0x78, 0x78, //1951
              0x96, 0xA5, 0xA6, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x78, 0x78, 0x87, 0x87, //1952
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x79, 0x77, 0x87, //1953
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x78, 0x79, 0x78, 0x68, 0x78, 0x87, //1954
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x79, 0x69, 0x78, 0x77, //1955
              0x96, 0xA5, 0xA5, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x78, 0x78, 0x87, 0x87, //1956
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x79, 0x77, 0x87, //1957
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x69, 0x78, 0x87, //1958
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x79, 0x69, 0x78, 0x77, //1959
              0x96, 0xA4, 0xA5, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //1960
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x78, 0x87, 0x87, //1961
              0x96, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x69, 0x78, 0x87, //1962
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x79, 0x69, 0x78, 0x77, //1963
              0x96, 0xA4, 0xA5, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //1964
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x78, 0x87, 0x87, //1965
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x69, 0x78, 0x87, //1966
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x79, 0x69, 0x78, 0x77, //1967
              0x96, 0xA4, 0xA5, 0xA5, 0xA6, 0xA6, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //1968
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x78, 0x87, 0x87, //1969
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x69, 0x78, 0x87, //1970
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x79, 0x69, 0x78, 0x77, //1971
              0x96, 0xA4, 0xA5, 0xA5, 0xA6, 0xA6, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //1972
              0xA5, 0xB5, 0x96, 0xA5, 0xA6, 0x96, 0x88, 0x78, 0x78, 0x78, 0x87, 0x87, //1973
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x69, 0x78, 0x87, //1974
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x78, 0x69, 0x78, 0x77, //1975
              0x96, 0xA4, 0xA5, 0xB5, 0xA6, 0xA6, 0x88, 0x89, 0x88, 0x78, 0x87, 0x87, //1976
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x88, 0x78, 0x78, 0x87, 0x87, //1977
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x79, 0x78, 0x87, //1978
              0x96, 0xB4, 0x96, 0xA6, 0x96, 0x97, 0x78, 0x79, 0x78, 0x69, 0x78, 0x77, //1979
              0x96, 0xA4, 0xA5, 0xB5, 0xA6, 0xA6, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //1980
              0xA5, 0xB4, 0x96, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x78, 0x78, 0x77, 0x87, //1981
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x79, 0x77, 0x87, //1982
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x78, 0x79, 0x78, 0x69, 0x78, 0x77, //1983
              0x96, 0xB4, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x88, 0x78, 0x87, 0x87, //1984
              0xA5, 0xB4, 0xA6, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x78, 0x78, 0x87, 0x87, //1985
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x79, 0x77, 0x87, //1986
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x79, 0x78, 0x69, 0x78, 0x87, //1987
              0x96, 0xB4, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x88, 0x78, 0x87, 0x86, //1988
              0xA5, 0xB4, 0xA5, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //1989
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x79, 0x77, 0x87, //1990
              0x95, 0xB4, 0x96, 0xA5, 0x86, 0x97, 0x88, 0x78, 0x78, 0x69, 0x78, 0x87, //1991
              0x96, 0xB4, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x88, 0x78, 0x87, 0x86, //1992
              0xA5, 0xB3, 0xA5, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //1993
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x78, 0x87, 0x87, //1994
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x76, 0x78, 0x69, 0x78, 0x87, //1995
              0x96, 0xB4, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x88, 0x78, 0x87, 0x86, //1996
              0xA5, 0xB3, 0xA5, 0xA5, 0xA6, 0xA6, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //1997
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x78, 0x87, 0x87, //1998
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x69, 0x78, 0x87, //1999
              0x96, 0xB4, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x88, 0x78, 0x87, 0x86, //2000
              0xA5, 0xB3, 0xA5, 0xA5, 0xA6, 0xA6, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //2001
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x78, 0x87, 0x87, //2002
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x69, 0x78, 0x87, //2003
              0x96, 0xB4, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x88, 0x78, 0x87, 0x86, //2004
              0xA5, 0xB3, 0xA5, 0xA5, 0xA6, 0xA6, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //2005
              0xA5, 0xB4, 0x96, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x78, 0x78, 0x87, 0x87, //2006
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x69, 0x78, 0x87, //2007
              0x96, 0xB4, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x87, 0x78, 0x87, 0x86, //2008
              0xA5, 0xB3, 0xA5, 0xB5, 0xA6, 0xA6, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //2009
              0xA5, 0xB4, 0x96, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x78, 0x78, 0x87, 0x87, //2010
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x79, 0x78, 0x87, //2011
              0x96, 0xB4, 0xA5, 0xB5, 0xA5, 0xA6, 0x87, 0x88, 0x87, 0x78, 0x87, 0x86, //2012
              0xA5, 0xB3, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x88, 0x78, 0x87, 0x87, //2013
              0xA5, 0xB4, 0x96, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x78, 0x78, 0x87, 0x87, //2014
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x79, 0x77, 0x87, //2015
              0x95, 0xB4, 0xA5, 0xB4, 0xA5, 0xA6, 0x87, 0x88, 0x87, 0x78, 0x87, 0x86, //2016
              0xA5, 0xC3, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x88, 0x78, 0x87, 0x87, //2017
              0xA5, 0xB4, 0xA6, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x78, 0x78, 0x87, 0x87, //2018
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x79, 0x77, 0x87, //2019
              0x95, 0xB4, 0xA5, 0xB4, 0xA5, 0xA6, 0x97, 0x87, 0x87, 0x78, 0x87, 0x86, //2020
              0xA5, 0xC3, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x88, 0x78, 0x87, 0x86, //2021
              0xA5, 0xB4, 0xA5, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //2022
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x79, 0x77, 0x87, //2023
              0x95, 0xB4, 0xA5, 0xB4, 0xA5, 0xA6, 0x97, 0x87, 0x87, 0x78, 0x87, 0x96, //2024
              0xA5, 0xC3, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x88, 0x78, 0x87, 0x86, //2025
              0xA5, 0xB3, 0xA5, 0xA5, 0xA6, 0xA6, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //2026
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x78, 0x87, 0x87, //2027
              0x95, 0xB4, 0xA5, 0xB4, 0xA5, 0xA6, 0x97, 0x87, 0x87, 0x78, 0x87, 0x96, //2028
              0xA5, 0xC3, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x88, 0x78, 0x87, 0x86, //2029
              0xA5, 0xB3, 0xA5, 0xA5, 0xA6, 0xA6, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //2030
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x78, 0x87, 0x87, //2031
              0x95, 0xB4, 0xA5, 0xB4, 0xA5, 0xA6, 0x97, 0x87, 0x87, 0x78, 0x87, 0x96, //2032
              0xA5, 0xC3, 0xA5, 0xB5, 0xA6, 0xA6, 0x88, 0x88, 0x88, 0x78, 0x87, 0x86, //2033
              0xA5, 0xB3, 0xA5, 0xA5, 0xA6, 0xA6, 0x88, 0x78, 0x88, 0x78, 0x87, 0x87, //2034
              0xA5, 0xB4, 0x96, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x78, 0x78, 0x87, 0x87, //2035
              0x95, 0xB4, 0xA5, 0xB4, 0xA5, 0xA6, 0x97, 0x87, 0x87, 0x78, 0x87, 0x96, //2036
              0xA5, 0xC3, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x88, 0x78, 0x87, 0x86, //2037
              0xA5, 0xB3, 0xA5, 0xA5, 0xA6, 0xA6, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //2038
              0xA5, 0xB4, 0x96, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x78, 0x78, 0x87, 0x87, //2039
              0x95, 0xB4, 0xA5, 0xB4, 0xA5, 0xA6, 0x97, 0x87, 0x87, 0x78, 0x87, 0x96, //2040
              0xA5, 0xC3, 0xA5, 0xB5, 0xA5, 0xA6, 0x87, 0x88, 0x87, 0x78, 0x87, 0x86, //2041
              0xA5, 0xB3, 0xA5, 0xB5, 0xA6, 0xA6, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //2042
              0xA5, 0xB4, 0x96, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x78, 0x78, 0x87, 0x87, //2043
              0x95, 0xB4, 0xA5, 0xB4, 0xA5, 0xA6, 0x97, 0x87, 0x87, 0x88, 0x87, 0x96, //2044
              0xA5, 0xC3, 0xA5, 0xB4, 0xA5, 0xA6, 0x87, 0x88, 0x87, 0x78, 0x87, 0x86, //2045
              0xA5, 0xB3, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x88, 0x78, 0x87, 0x87, //2046
              0xA5, 0xB4, 0x96, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x78, 0x78, 0x87, 0x87, //2047
              0x95, 0xB4, 0xA5, 0xB4, 0xA5, 0xA5, 0x97, 0x87, 0x87, 0x88, 0x86, 0x96, //2048
              0xA4, 0xC3, 0xA5, 0xA5, 0xA5, 0xA6, 0x97, 0x87, 0x87, 0x78, 0x87, 0x86, //2049
              0xA5, 0xC3, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x78, 0x78, 0x87, 0x87}; //2050

            // 位元陣列 gLanarHoliDay 存放每年的二十四節氣對應的陽曆日期
            // 每年的二十四節氣對應的陽曆日期幾乎固定，平均分佈於十二個月中
            // 1月 2月 3月 4月 5月 6月
            // 小寒 大寒 立春 雨水 驚蟄 春分 清明 谷雨 立夏 小滿 芒種 夏至
            // 7月 8月 9月 10月 11月 12月
            // 小暑 大暑 立秋 處暑 白露 秋分 寒露 霜降 立冬 小雪 大雪 冬至
            string[] LunarHolDayName =
                  {
                  "小寒", "大寒", "立春", "雨水",
                  "驚蟄", "春分", "清明", "穀雨",
                  "立夏", "小滿", "芒種", "夏至",
                  "小暑", "大暑", "立秋", "處暑",
                  "白露", "秋分", "寒露", "霜降",
                  "立冬", "小雪", "大雪", "冬至"};

            const ushort START_YEAR = 1901;
            const ushort END_YEAR = 2050;

            byte Flag;
            int Day, iYear, iMonth, iDay;
            iYear = date1.Year;
            if ((iYear < 1901) || (iYear > END_YEAR))
            {
                return "";
            }

            iMonth = date1.Month;
            iDay = date1.Day;
            Flag = gLunarHoliDay[(iYear - START_YEAR) * 12 + iMonth - 1];

            if (iDay < 15)
            {
                Day = 15 - ((Flag >> 4) & 0x0f);
            }
            else
            {
                Day = (Flag & 0x0f) + 15;
            }
            string lunar24 = "";
            if (iDay == Day)
            {
                if (iDay > 15)
                {
                    lunar24 = LunarHolDayName[(iMonth - 1) * 2 + 1];
                }
                else
                {
                    lunar24 = LunarHolDayName[(iMonth - 1) * 2];
                }
            }
            else
            {
                if (iDay > 15)
                {
                    lunar24 = LunarHolDayName[(iMonth - 1) * 2 + 1];
                }
                else
                {
                    lunar24 = LunarHolDayName[(iMonth - 1) * 2];
                }
                //return "";
            }
            int seg24 = 0;
            switch (lunar24)
            {
                case "小寒": seg24 = 23; break;
                case "大寒": seg24 = 24; break;
                case "立春": seg24 = 1; break;
                case "雨水": seg24 = 2; break;
                case "驚蟄": seg24 = 3; break;
                case "春分": seg24 = 4; break;
                case "清明": seg24 = 5; break;
                case "穀雨": seg24 = 6; break;
                case "立夏": seg24 = 7; break;
                case "小滿": seg24 = 8; break;
                case "芒種": seg24 = 9; break;
                case "夏至": seg24 = 10; break;
                case "小暑": seg24 = 11; break;
                case "大暑": seg24 = 12; break;
                case "立秋": seg24 = 13; break;
                case "處暑": seg24 = 14; break;
                case "白露": seg24 = 15; break;
                case "秋分": seg24 = 16; break;
                case "寒露": seg24 = 17; break;
                case "霜降": seg24 = 18; break;
                case "立冬": seg24 = 19; break;
                case "小雪": seg24 = 20; break;
                case "大雪": seg24 = 21; break;
                case "冬至": seg24 = 22; break;
            }
            int mh = seg24 / 2;//结果为商1
            int d = seg24 % 2;//结果为余数
            if (d > 0)
            {
                mh = mh + 1;
                if (mh == 13)
                {
                    mh = 1;
                }
            }
            return mh.ToString();
        }//function GetLunarHolDay(DateTime date1)

        public string GetLunarHolDay(DateTime date1)
        {
            //*********************************************************************************
            // 節氣無任何確定規律,所以只好建立表格對應
            //**********************************************************************************}
            // 數據格式說明:
            // 如 1901 年的節氣為
            // 1月 2月 3月 4月 5月 6月 7月 8月 9月 10月 11月 12月
            // 6, 21, 4, 19, 6, 21, 5, 21, 6,22, 6,22, 8, 23, 8, 24, 8, 24, 8, 24, 8, 23, 8, 22
            // 9, 6, 11,4, 9, 6, 10,6, 9,7, 9,7, 7, 8, 7, 9, 7, 9, 7, 9, 7, 8, 7, 15
            // 上面第一行數據為每月節氣對應日期,15減去每月第一個節氣,每月第二個節氣減去15得第二行
            // 這樣每月兩個節氣對應數據都小於16,每月用一個陣列元素存放
            // 高位元(high-bit)存放第一個節氣數據,低位元(low-bit)存放第二個節氣的數據,可得下表
            byte[] gLunarHoliDay ={
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x79, 0x69, 0x78, 0x77, //1901
              0x96, 0xA4, 0x96, 0x96, 0x97, 0x87, 0x79, 0x79, 0x79, 0x69, 0x78, 0x78, //1902
              0x96, 0xA5, 0x87, 0x96, 0x87, 0x87, 0x79, 0x69, 0x69, 0x69, 0x78, 0x78, //1903
              0x86, 0xA5, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x79, 0x78, 0x87, //1904
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x79, 0x69, 0x78, 0x77, //1905
              0x96, 0xA4, 0x96, 0x96, 0x97, 0x97, 0x79, 0x79, 0x79, 0x69, 0x78, 0x78, //1906
              0x96, 0xA5, 0x87, 0x96, 0x87, 0x87, 0x79, 0x69, 0x69, 0x69, 0x78, 0x78, //1907
              0x86, 0xA5, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x69, 0x78, 0x87, //1908
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x79, 0x69, 0x78, 0x77, //1909
              0x96, 0xA4, 0x96, 0x96, 0x97, 0x97, 0x79, 0x79, 0x79, 0x69, 0x78, 0x78, //1910
              0x96, 0xA5, 0x87, 0x96, 0x87, 0x87, 0x79, 0x69, 0x69, 0x69, 0x78, 0x78, //1911
              0x86, 0xA5, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x69, 0x78, 0x87, //1912
              0x95, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x79, 0x69, 0x78, 0x77, //1913
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x79, 0x79, 0x79, 0x69, 0x78, 0x78, //1914
              0x96, 0xA5, 0x97, 0x96, 0x97, 0x87, 0x79, 0x79, 0x69, 0x69, 0x78, 0x78, //1915
              0x96, 0xA5, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x79, 0x77, 0x87, //1916
              0x95, 0xB4, 0x96, 0xA6, 0x96, 0x97, 0x78, 0x79, 0x78, 0x69, 0x78, 0x87, //1917
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x79, 0x79, 0x79, 0x69, 0x78, 0x77, //1918
              0x96, 0xA5, 0x97, 0x96, 0x97, 0x87, 0x79, 0x79, 0x69, 0x69, 0x78, 0x78, //1919
              0x96, 0xA5, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x79, 0x77, 0x87, //1920
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x78, 0x79, 0x78, 0x69, 0x78, 0x87, //1921
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x79, 0x79, 0x79, 0x69, 0x78, 0x77, //1922
              0x96, 0xA4, 0x96, 0x96, 0x97, 0x87, 0x79, 0x79, 0x69, 0x69, 0x78, 0x78, //1923
              0x96, 0xA5, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x79, 0x77, 0x87, //1924
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x78, 0x79, 0x78, 0x69, 0x78, 0x87, //1925
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x79, 0x69, 0x78, 0x77, //1926
              0x96, 0xA4, 0x96, 0x96, 0x97, 0x87, 0x79, 0x79, 0x79, 0x69, 0x78, 0x78, //1927
              0x96, 0xA5, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x78, 0x87, 0x87, //1928
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x79, 0x77, 0x87, //1929
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x79, 0x69, 0x78, 0x77, //1930
              0x96, 0xA4, 0x96, 0x96, 0x97, 0x87, 0x79, 0x79, 0x79, 0x69, 0x78, 0x78, //1931
              0x96, 0xA5, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x78, 0x87, 0x87, //1932
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x69, 0x78, 0x87, //1933
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x79, 0x69, 0x78, 0x77, //1934
              0x96, 0xA4, 0x96, 0x96, 0x97, 0x97, 0x79, 0x79, 0x79, 0x69, 0x78, 0x78, //1935
              0x96, 0xA5, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x78, 0x87, 0x87, //1936
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x69, 0x78, 0x87, //1937
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x79, 0x69, 0x78, 0x77, //1938
              0x96, 0xA4, 0x96, 0x96, 0x97, 0x97, 0x79, 0x79, 0x79, 0x69, 0x78, 0x78, //1939
              0x96, 0xA5, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x78, 0x87, 0x87, //1940
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x69, 0x78, 0x87, //1941
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x79, 0x69, 0x78, 0x77, //1942
              0x96, 0xA4, 0x96, 0x96, 0x97, 0x97, 0x79, 0x79, 0x79, 0x69, 0x78, 0x78, //1943
              0x96, 0xA5, 0x96, 0xA5, 0xA6, 0x96, 0x88, 0x78, 0x78, 0x78, 0x87, 0x87, //1944
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x79, 0x77, 0x87, //1945
              0x95, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x78, 0x69, 0x78, 0x77, //1946
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x79, 0x79, 0x79, 0x69, 0x78, 0x78, //1947
              0x96, 0xA5, 0xA6, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x78, 0x78, 0x87, 0x87, //1948
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x79, 0x78, 0x79, 0x77, 0x87, //1949
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x78, 0x79, 0x78, 0x69, 0x78, 0x77, //1950
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x79, 0x79, 0x79, 0x69, 0x78, 0x78, //1951
              0x96, 0xA5, 0xA6, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x78, 0x78, 0x87, 0x87, //1952
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x79, 0x77, 0x87, //1953
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x78, 0x79, 0x78, 0x68, 0x78, 0x87, //1954
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x79, 0x69, 0x78, 0x77, //1955
              0x96, 0xA5, 0xA5, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x78, 0x78, 0x87, 0x87, //1956
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x79, 0x77, 0x87, //1957
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x69, 0x78, 0x87, //1958
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x79, 0x69, 0x78, 0x77, //1959
              0x96, 0xA4, 0xA5, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //1960
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x78, 0x87, 0x87, //1961
              0x96, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x69, 0x78, 0x87, //1962
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x79, 0x69, 0x78, 0x77, //1963
              0x96, 0xA4, 0xA5, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //1964
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x78, 0x87, 0x87, //1965
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x69, 0x78, 0x87, //1966
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x79, 0x69, 0x78, 0x77, //1967
              0x96, 0xA4, 0xA5, 0xA5, 0xA6, 0xA6, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //1968
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x78, 0x87, 0x87, //1969
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x69, 0x78, 0x87, //1970
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x79, 0x69, 0x78, 0x77, //1971
              0x96, 0xA4, 0xA5, 0xA5, 0xA6, 0xA6, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //1972
              0xA5, 0xB5, 0x96, 0xA5, 0xA6, 0x96, 0x88, 0x78, 0x78, 0x78, 0x87, 0x87, //1973
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x69, 0x78, 0x87, //1974
              0x96, 0xB4, 0x96, 0xA6, 0x97, 0x97, 0x78, 0x79, 0x78, 0x69, 0x78, 0x77, //1975
              0x96, 0xA4, 0xA5, 0xB5, 0xA6, 0xA6, 0x88, 0x89, 0x88, 0x78, 0x87, 0x87, //1976
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x88, 0x78, 0x78, 0x87, 0x87, //1977
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x79, 0x78, 0x87, //1978
              0x96, 0xB4, 0x96, 0xA6, 0x96, 0x97, 0x78, 0x79, 0x78, 0x69, 0x78, 0x77, //1979
              0x96, 0xA4, 0xA5, 0xB5, 0xA6, 0xA6, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //1980
              0xA5, 0xB4, 0x96, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x78, 0x78, 0x77, 0x87, //1981
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x79, 0x77, 0x87, //1982
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x78, 0x79, 0x78, 0x69, 0x78, 0x77, //1983
              0x96, 0xB4, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x88, 0x78, 0x87, 0x87, //1984
              0xA5, 0xB4, 0xA6, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x78, 0x78, 0x87, 0x87, //1985
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x79, 0x77, 0x87, //1986
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x79, 0x78, 0x69, 0x78, 0x87, //1987
              0x96, 0xB4, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x88, 0x78, 0x87, 0x86, //1988
              0xA5, 0xB4, 0xA5, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //1989
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x79, 0x77, 0x87, //1990
              0x95, 0xB4, 0x96, 0xA5, 0x86, 0x97, 0x88, 0x78, 0x78, 0x69, 0x78, 0x87, //1991
              0x96, 0xB4, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x88, 0x78, 0x87, 0x86, //1992
              0xA5, 0xB3, 0xA5, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //1993
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x78, 0x87, 0x87, //1994
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x76, 0x78, 0x69, 0x78, 0x87, //1995
              0x96, 0xB4, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x88, 0x78, 0x87, 0x86, //1996
              0xA5, 0xB3, 0xA5, 0xA5, 0xA6, 0xA6, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //1997
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x78, 0x87, 0x87, //1998
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x69, 0x78, 0x87, //1999
              0x96, 0xB4, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x88, 0x78, 0x87, 0x86, //2000
              0xA5, 0xB3, 0xA5, 0xA5, 0xA6, 0xA6, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //2001
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x78, 0x87, 0x87, //2002
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x69, 0x78, 0x87, //2003
              0x96, 0xB4, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x88, 0x78, 0x87, 0x86, //2004
              0xA5, 0xB3, 0xA5, 0xA5, 0xA6, 0xA6, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //2005
              0xA5, 0xB4, 0x96, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x78, 0x78, 0x87, 0x87, //2006
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x69, 0x78, 0x87, //2007
              0x96, 0xB4, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x87, 0x78, 0x87, 0x86, //2008
              0xA5, 0xB3, 0xA5, 0xB5, 0xA6, 0xA6, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //2009
              0xA5, 0xB4, 0x96, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x78, 0x78, 0x87, 0x87, //2010
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x79, 0x78, 0x87, //2011
              0x96, 0xB4, 0xA5, 0xB5, 0xA5, 0xA6, 0x87, 0x88, 0x87, 0x78, 0x87, 0x86, //2012
              0xA5, 0xB3, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x88, 0x78, 0x87, 0x87, //2013
              0xA5, 0xB4, 0x96, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x78, 0x78, 0x87, 0x87, //2014
              0x95, 0xB4, 0x96, 0xA5, 0x96, 0x97, 0x88, 0x78, 0x78, 0x79, 0x77, 0x87, //2015
              0x95, 0xB4, 0xA5, 0xB4, 0xA5, 0xA6, 0x87, 0x88, 0x87, 0x78, 0x87, 0x86, //2016
              0xA5, 0xC3, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x88, 0x78, 0x87, 0x87, //2017
              0xA5, 0xB4, 0xA6, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x78, 0x78, 0x87, 0x87, //2018
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x79, 0x77, 0x87, //2019
              0x95, 0xB4, 0xA5, 0xB4, 0xA5, 0xA6, 0x97, 0x87, 0x87, 0x78, 0x87, 0x86, //2020
              0xA5, 0xC3, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x88, 0x78, 0x87, 0x86, //2021
              0xA5, 0xB4, 0xA5, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //2022
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x79, 0x77, 0x87, //2023
              0x95, 0xB4, 0xA5, 0xB4, 0xA5, 0xA6, 0x97, 0x87, 0x87, 0x78, 0x87, 0x96, //2024
              0xA5, 0xC3, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x88, 0x78, 0x87, 0x86, //2025
              0xA5, 0xB3, 0xA5, 0xA5, 0xA6, 0xA6, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //2026
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x78, 0x87, 0x87, //2027
              0x95, 0xB4, 0xA5, 0xB4, 0xA5, 0xA6, 0x97, 0x87, 0x87, 0x78, 0x87, 0x96, //2028
              0xA5, 0xC3, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x88, 0x78, 0x87, 0x86, //2029
              0xA5, 0xB3, 0xA5, 0xA5, 0xA6, 0xA6, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //2030
              0xA5, 0xB4, 0x96, 0xA5, 0x96, 0x96, 0x88, 0x78, 0x78, 0x78, 0x87, 0x87, //2031
              0x95, 0xB4, 0xA5, 0xB4, 0xA5, 0xA6, 0x97, 0x87, 0x87, 0x78, 0x87, 0x96, //2032
              0xA5, 0xC3, 0xA5, 0xB5, 0xA6, 0xA6, 0x88, 0x88, 0x88, 0x78, 0x87, 0x86, //2033
              0xA5, 0xB3, 0xA5, 0xA5, 0xA6, 0xA6, 0x88, 0x78, 0x88, 0x78, 0x87, 0x87, //2034
              0xA5, 0xB4, 0x96, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x78, 0x78, 0x87, 0x87, //2035
              0x95, 0xB4, 0xA5, 0xB4, 0xA5, 0xA6, 0x97, 0x87, 0x87, 0x78, 0x87, 0x96, //2036
              0xA5, 0xC3, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x88, 0x78, 0x87, 0x86, //2037
              0xA5, 0xB3, 0xA5, 0xA5, 0xA6, 0xA6, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //2038
              0xA5, 0xB4, 0x96, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x78, 0x78, 0x87, 0x87, //2039
              0x95, 0xB4, 0xA5, 0xB4, 0xA5, 0xA6, 0x97, 0x87, 0x87, 0x78, 0x87, 0x96, //2040
              0xA5, 0xC3, 0xA5, 0xB5, 0xA5, 0xA6, 0x87, 0x88, 0x87, 0x78, 0x87, 0x86, //2041
              0xA5, 0xB3, 0xA5, 0xB5, 0xA6, 0xA6, 0x88, 0x88, 0x88, 0x78, 0x87, 0x87, //2042
              0xA5, 0xB4, 0x96, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x78, 0x78, 0x87, 0x87, //2043
              0x95, 0xB4, 0xA5, 0xB4, 0xA5, 0xA6, 0x97, 0x87, 0x87, 0x88, 0x87, 0x96, //2044
              0xA5, 0xC3, 0xA5, 0xB4, 0xA5, 0xA6, 0x87, 0x88, 0x87, 0x78, 0x87, 0x86, //2045
              0xA5, 0xB3, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x88, 0x78, 0x87, 0x87, //2046
              0xA5, 0xB4, 0x96, 0xA5, 0xA6, 0x96, 0x88, 0x88, 0x78, 0x78, 0x87, 0x87, //2047
              0x95, 0xB4, 0xA5, 0xB4, 0xA5, 0xA5, 0x97, 0x87, 0x87, 0x88, 0x86, 0x96, //2048
              0xA4, 0xC3, 0xA5, 0xA5, 0xA5, 0xA6, 0x97, 0x87, 0x87, 0x78, 0x87, 0x86, //2049
              0xA5, 0xC3, 0xA5, 0xB5, 0xA6, 0xA6, 0x87, 0x88, 0x78, 0x78, 0x87, 0x87}; //2050

            // 位元陣列 gLanarHoliDay 存放每年的二十四節氣對應的陽曆日期
            // 每年的二十四節氣對應的陽曆日期幾乎固定，平均分佈於十二個月中
            // 1月 2月 3月 4月 5月 6月
            // 小寒 大寒 立春 雨水 驚蟄 春分 清明 谷雨 立夏 小滿 芒種 夏至
            // 7月 8月 9月 10月 11月 12月
            // 小暑 大暑 立秋 處暑 白露 秋分 寒露 霜降 立冬 小雪 大雪 冬至
            string[] LunarHolDayName =
                  {
                  "小寒", "大寒", "立春", "雨水",
                  "驚蟄", "春分", "清明", "穀雨",
                  "立夏", "小滿", "芒種", "夏至",
                  "小暑", "大暑", "立秋", "處暑",
                  "白露", "秋分", "寒露", "霜降",
                  "立冬", "小雪", "大雪", "冬至"};

            const ushort START_YEAR = 1901;
            const ushort END_YEAR = 2050;

            byte Flag;
            int Day, iYear, iMonth, iDay;
            iYear = date1.Year;
            if ((iYear < 1901) || (iYear > END_YEAR))
            {
                return "";
            }

            iMonth = date1.Month;
            iDay = date1.Day;
            Flag = gLunarHoliDay[(iYear - START_YEAR) * 12 + iMonth - 1];

            if (iDay < 15)
            {
                Day = 15 - ((Flag >> 4) & 0x0f);
            }
            else
            {
                Day = (Flag & 0x0f) + 15;
            }
            string lunar24 = "";
            if (iDay == Day)
            {
                if (iDay > 15)
                {
                    lunar24= LunarHolDayName[(iMonth - 1) * 2 + 1];
                }
                else
                {
                    lunar24= LunarHolDayName[(iMonth - 1) * 2];
                }
            }
            else
            {
                if (iDay > 15)
                {
                    lunar24= LunarHolDayName[(iMonth - 1) * 2 + 1];
                }
                else
                {
                    lunar24= LunarHolDayName[(iMonth - 1) * 2];
                }
                //return "";
            }
            int seg24 = 0;
            switch (lunar24)
            {
                case "小寒": seg24 = 23; break;
                case "大寒": seg24 = 24; break;
                case "立春": seg24 = 1; break;
                case "雨水": seg24 = 2; break;
                case "驚蟄": seg24 = 3; break;
                case "春分": seg24 = 4; break;
                case "清明": seg24 = 5; break;
                case "穀雨": seg24 = 6; break;
                case "立夏": seg24 = 7; break;
                case "小滿": seg24 = 8; break;
                case "芒種": seg24 = 9; break;
                case "夏至": seg24 = 10; break;
                case "小暑": seg24 = 11; break;
                case "大暑": seg24 = 12; break;
                case "立秋": seg24 = 13; break;
                case "處暑": seg24 = 14; break;
                case "白露": seg24 = 15; break;
                case "秋分": seg24 = 16; break;
                case "寒露": seg24 = 17; break;
                case "霜降": seg24 = 18; break;
                case "立冬": seg24 = 19; break;
                case "小雪": seg24 = 20; break;
                case "大雪": seg24 = 21; break;
                case "冬至": seg24 = 22; break;
            }
            int mh = seg24 / 2;//结果为商1
            int d = seg24 % 2;//结果为余数
            if (d > 0)
            {
                mh = mh + 1;
                if (mh == 13)
                {
                    mh = 1;
                }
            }
            return mh.ToString();
         }//function GetLunarHolDay(DateTime date1)

        #region 按公??算的?日
        private static SolarHolidayStruct[] sHolidayInfo = new SolarHolidayStruct[]{
            new SolarHolidayStruct(1, 1, 1, "元旦"),
            new SolarHolidayStruct(2, 2, 0, "世界?地日"),
            new SolarHolidayStruct(2, 10, 0, "??氣象?"),
            new SolarHolidayStruct(2, 14, 0, "情人?"),
            new SolarHolidayStruct(3, 1, 0, "??海豹日"),
            new SolarHolidayStruct(3, 5, 0, "?雷??念日"),
            new SolarHolidayStruct(3, 8, 0, "?女?"), 
            new SolarHolidayStruct(3, 12, 0, "植?? ?中山逝世?念日"), 
            new SolarHolidayStruct(3, 14, 0, "??員警日"),
            new SolarHolidayStruct(3, 15, 0, "消?者?益日"),
            new SolarHolidayStruct(3, 17, 0, "中???? ??航海日"),
            new SolarHolidayStruct(3, 21, 0, "世界森林日 消除種族歧???日 世界兒歌日"),
            new SolarHolidayStruct(3, 22, 0, "世界水日"),
            new SolarHolidayStruct(3, 24, 0, "世界防治?核病日"),
            new SolarHolidayStruct(4, 1, 0, "愚人?"),
            new SolarHolidayStruct(4, 7, 0, "世界?生日"),
            new SolarHolidayStruct(4, 22, 0, "世界地球日"),
            new SolarHolidayStruct(5, 1, 1, "???"), 
            new SolarHolidayStruct(5, 2, 1, "???假日"),
            new SolarHolidayStruct(5, 3, 1, "???假日"),
            new SolarHolidayStruct(5, 4, 0, "青年?"), 
            new SolarHolidayStruct(5, 8, 0, "世界?十字日"),
            new SolarHolidayStruct(5, 12, 0, "???士?"), 
            new SolarHolidayStruct(5, 31, 0, "世界??日"), 
            new SolarHolidayStruct(6, 1, 0, "??兒童?"), 
            new SolarHolidayStruct(6, 5, 0, "世界?境保?日"),
            new SolarHolidayStruct(6, 26, 0, "??禁毒日"),
            new SolarHolidayStruct(7, 1, 0, "建黨? 香港回??念 世界建築日"),
            new SolarHolidayStruct(7, 11, 0, "世界人口日"),
            new SolarHolidayStruct(8, 1, 0, "建??"), 
            new SolarHolidayStruct(8, 8, 0, "中?男子? 父??"),
            new SolarHolidayStruct(8, 15, 0, "抗日??勝利?念"),
            new SolarHolidayStruct(9, 9, 0, "毛主席逝世?念"), 
            new SolarHolidayStruct(9, 10, 0, "教??"), 
            new SolarHolidayStruct(9, 18, 0, "九‧一八事??念日"),
            new SolarHolidayStruct(9, 20, 0, "???牙日"),
            new SolarHolidayStruct(9, 27, 0, "世界旅遊日"),
            new SolarHolidayStruct(9, 28, 0, "孔子?辰"),
            new SolarHolidayStruct(10, 1, 1, "??? ??音?日"),
            new SolarHolidayStruct(10, 2, 1, "???假日"),
            new SolarHolidayStruct(10, 3, 1, "???假日"),
            new SolarHolidayStruct(10, 6, 0, "老人?"), 
            new SolarHolidayStruct(10, 24, 0, "?合?日"),
            new SolarHolidayStruct(11, 10, 0, "世界青年?"),
            new SolarHolidayStruct(11, 12, 0, "?中山?辰?念"), 
            new SolarHolidayStruct(12, 1, 0, "世界愛滋病日"), 
            new SolarHolidayStruct(12, 3, 0, "世界?疾人日"), 
            new SolarHolidayStruct(12, 20, 0, "澳?回??念"), 
            new SolarHolidayStruct(12, 24, 0, "平安夜"), 
            new SolarHolidayStruct(12, 25, 0, "聖??"), 
            new SolarHolidayStruct(12, 26, 0, "毛主席?辰?念")
           };
        #endregion

        #region 按???算的?日
        private static LunarHolidayStruct[] lHolidayInfo = new LunarHolidayStruct[]{
            new LunarHolidayStruct(1, 1, 1, "春節"), 
            new LunarHolidayStruct(1, 15, 0, "元宵節"), 
            new LunarHolidayStruct(5, 5, 0, "端午節"), 
            new LunarHolidayStruct(7, 7, 0, "七夕情人?"),
            new LunarHolidayStruct(7, 15, 0, "中元節"), 
            new LunarHolidayStruct(8, 15, 0, "中秋節"), 
            new LunarHolidayStruct(9, 9, 0, "重陽節"), 
            new LunarHolidayStruct(12, 8, 0, "臘八節"),
            new LunarHolidayStruct(12, 23, 0, "北方小年"),
            new LunarHolidayStruct(12, 24, 0, "南方小年"),
            //new LunarHolidayStruct(12, 30, 0, "除夕")  //注意除夕需要其它方法?行?算
        };
        #endregion

        #region 按某月第幾?星期幾
        private static WeekHolidayStruct[] wHolidayInfo = new WeekHolidayStruct[]{
            new WeekHolidayStruct(5, 2, 1, "母??"), 
            new WeekHolidayStruct(5, 3, 1, "全?助?日"), 
            new WeekHolidayStruct(6, 3, 1, "父??"), 
            new WeekHolidayStruct(9, 3, 3, "??和平日"), 
            new WeekHolidayStruct(9, 4, 1, "???人?"), 
            new WeekHolidayStruct(10, 1, 2, "??住房日"), 
            new WeekHolidayStruct(10, 1, 4, "????自然?害日"),
            new WeekHolidayStruct(11, 4, 5, "感恩?")
        };
        #endregion

        #endregion

        #region 構造函?
        #region ChinaCalendar <公?日期初始化>
        /// <summary>
        /// 用一??准的公?日期?初使化
        /// </summary>
        /// <param name="dt"></param>
        public EcanChineseCalendar(DateTime dt)
        {
            int i;
            int leap;
            int temp;
            int offset;

            CheckDateLimit(dt);

            _date = dt.Date;
            _datetime = dt;

            //??日期?算部分
            leap = 0;
            temp = 0;

            TimeSpan ts = _date - EcanChineseCalendar.MinDay;//?算?天的基本差距
            offset = ts.Days;

            for (i = MinYear; i <= MaxYear; i++)
            {
                temp = GetChineseYearDays(i);  //求?年??年天?
                if (offset - temp < 1)
                    break;
                else
                {
                    offset = offset - temp;
                }
            }
            _cYear = i;

            leap = GetChineseLeapMonth(_cYear);//?算?年?哪?月
            //?定?年是否有?月
            if (leap > 0)
            {
                _cIsLeapYear = true;
            }
            else
            {
                _cIsLeapYear = false;
            }

            _cIsLeapMonth = false;
            for (i = 1; i <= 12; i++)
            {
                //?月
                if ((leap > 0) && (i == leap + 1) && (_cIsLeapMonth == false))
                {
                    _cIsLeapMonth = true;
                    i = i - 1;
                    temp = GetChineseLeapMonthDays(_cYear); //?算?月天?
                }
                else
                {
                    _cIsLeapMonth = false;
                    temp = GetChineseMonthDays(_cYear, i);//?算非?月天?
                }

                offset = offset - temp;
                if (offset <= 0) break;
            }

            offset = offset + temp;
            _cMonth = i;
            _cDay = offset;
        }
        #endregion

        #region ChinaCalendar <??日期初始化>
        /// <summary>
        /// 用??的日期?初使化
        /// </summary>
        /// <param name="cy">??年</param>
        /// <param name="cm">??月</param>
        /// <param name="cd">??日</param>
        /// <param name="LeapFlag">?月?志</param>
        public EcanChineseCalendar(int cy, int cm, int cd, bool leapMonthFlag)
        {
            int i, leap, Temp, offset;

            CheckChineseDateLimit(cy, cm, cd, leapMonthFlag);

            _cYear = cy;
            _cMonth = cm;
            _cDay = cd;

            offset = 0;

            for (i = MinYear; i < cy; i++)
            {
                Temp = GetChineseYearDays(i); //求?年??年天?
                offset = offset + Temp;
            }

            leap = GetChineseLeapMonth(cy);// ?算?年???哪?月
            if (leap != 0)
            {
                this._cIsLeapYear = true;
            }
            else
            {
                this._cIsLeapYear = false;
            }

            if (cm != leap)
            {
                _cIsLeapMonth = false;  //?前日期並非?月
            }
            else
            {
                _cIsLeapMonth = leapMonthFlag;  //使用用??入的是否?月月份
            }


            if ((_cIsLeapYear == false) || //?年?有?月
                 (cm < leap)) //?算月份小於?月     
            {
                #region ...
                for (i = 1; i < cm; i++)
                {
                    Temp = GetChineseMonthDays(cy, i);//?算非?月天?
                    offset = offset + Temp;
                }

                //?查日期是否大於最大天
                if (cd > GetChineseMonthDays(cy, cm))
                {
                    throw new newCalendarException("不合法的??日期");
                }
                offset = offset + cd; //加上?月的天?
                #endregion
            }
            else   //是?年，且?算月份大於或等於?月
            {
                #region ...
                for (i = 1; i < cm; i++)
                {
                    Temp = GetChineseMonthDays(cy, i); //?算非?月天?
                    offset = offset + Temp;
                }

                if (cm > leap) //?算月大於?月
                {
                    Temp = GetChineseLeapMonthDays(cy);   //?算?月天?
                    offset = offset + Temp;               //加上?月天?

                    if (cd > GetChineseMonthDays(cy, cm))
                    {
                        throw new newCalendarException("不合法的??日期");
                    }
                    offset = offset + cd;
                }
                else  //?算月等於?月
                {
                    //如果需要?算的是?月，??首先加上與?月??的普通月的天?
                    if (this._cIsLeapMonth == true) //?算月??月
                    {
                        Temp = GetChineseMonthDays(cy, cm); //?算非?月天?
                        offset = offset + Temp;
                    }

                    if (cd > GetChineseLeapMonthDays(cy))
                    {
                        throw new newCalendarException("不合法的??日期");
                    }
                    offset = offset + cd;
                }
                #endregion
            }


            _date = MinDay.AddDays(offset);
        }
        #endregion
        #endregion

        #region 私有函?

        #region GetChineseMonthDays
        //?回?? y年m月的?天?
        private int GetChineseMonthDays(int year, int month)
        {
            if (BitTest32((LunarDateArray[year - MinYear] & 0x0000FFFF), (16 - month)))
            {
                return 30;
            }
            else
            {
                return 29;
            }
        }
        #endregion

        #region GetChineseLeapMonth
        //?回?? y年?哪?月 1-12 , ???回 0
        private int GetChineseLeapMonth(int year)
        {

            return LunarDateArray[year - MinYear] & 0xF;

        }
        #endregion

        #region GetChineseLeapMonthDays
        //?回?? y年?月的天?
        private int GetChineseLeapMonthDays(int year)
        {
            if (GetChineseLeapMonth(year) != 0)
            {
                if ((LunarDateArray[year - MinYear] & 0x10000) != 0)
                {
                    return 30;
                }
                else
                {
                    return 29;
                }
            }
            else
            {
                return 0;
            }
        }
        #endregion

        #region GetChineseYearDays
        /// <summary>
        /// 取??年一年的天?
        /// </summary>
        /// <param name="year"></param>
        /// <returns></returns>
        private int GetChineseYearDays(int year)
        {
            int i, f, sumDay, info;

            sumDay = 348; //29天 X 12?月
            i = 0x8000;
            info = LunarDateArray[year - MinYear] & 0x0FFFF;

            //?算12?月中有多少天?30天
            for (int m = 0; m < 12; m++)
            {
                f = info & i;
                if (f != 0)
                {
                    sumDay++;
                }
                i = i >> 1;
            }
            return sumDay + GetChineseLeapMonthDays(year);
        }
        #endregion

        #region GetChineseHour
        /// <summary>
        /// ?得?前??的?辰
        /// </summary>
        /// <param name="time"></param>
        /// <returns></returns>
        /// 
        private string GetChineseHour(DateTime dt)
        {

            int _hour, _minute, offset, i;
            int indexGan;
            string ganHour, zhiHour;
            string tmpGan;

            //?算?辰的地支
            _hour = dt.Hour;    //?得?前??小?
            _minute = dt.Minute;  //?得?前??分?

            //if (_minute != 0) _hour += 1;
            _hour += 1;
            offset = _hour / 2;
            if (offset >= 12) offset = 0;
            //zhiHour = zhiStr[offset].ToString();

            //?算天干
            TimeSpan ts = this._date - GanZhiStartDay;
            i = ts.Days % 60;

            indexGan = ((i % 10 + 1) * 2 - 1) % 10 - 1; //ganStr[i % 10] ?日的天干,(n*2-1) %10得出地支??,n?1?始
            tmpGan = ganStr.Substring(indexGan) + ganStr.Substring(0, indexGan + 2);//??12位
            //ganHour = ganStr[((i % 10 + 1) * 2 - 1) % 10 - 1].ToString();
            if (_hour == 24 )
            {
                return tmpGan[offset+2].ToString() + zhiStr[offset].ToString();
            }
            else
            {
                return tmpGan[offset].ToString() + zhiStr[offset].ToString();
            }            

        }
        #endregion

        #region CheckDateLimit
        /// <summary>
        /// ?查公?日期是否符合要求
        /// </summary>
        /// <param name="dt"></param>
        private void CheckDateLimit(DateTime dt)
        {
            if ((dt < MinDay) || (dt > MaxDay))
            {
                throw new newCalendarException("超出可??的日期");
            }
        }
        #endregion

        #region CheckChineseDateLimit
        /// <summary>
        /// ?查??日期是否合理
        /// </summary>
        /// <param name="year"></param>
        /// <param name="month"></param>
        /// <param name="day"></param>
        /// <param name="leapMonth"></param>
        private void CheckChineseDateLimit(int year, int month, int day, bool leapMonth)
        {
            if ((year < MinYear) || (year > MaxYear))
            {
                throw new newCalendarException("非法??日期");
            }
            if ((month < 1) || (month > 12))
            {
                throw new newCalendarException("非法??日期");
            }
            if ((day < 1) || (day > 30)) //中?的月最多30天
            {
                throw new newCalendarException("非法??日期");
            }

            int leap = GetChineseLeapMonth(year);// ?算?年???哪?月
            if ((leapMonth == true) && (month != leap))
            {
                throw new newCalendarException("非法??日期");
            }


        }
        #endregion

        #region ConvertNumToChineseNum
        /// <summary>
        /// ?0-9?成?字形式
        /// </summary>
        /// <param name="n"></param>
        /// <returns></returns>
        private string ConvertNumToChineseNum(char n)
        {
            if ((n < '0') || (n > '9')) return "";
            switch (n)
            {
                case '0':
                    return HZNum[0].ToString();
                case '1':
                    return HZNum[1].ToString();
                case '2':
                    return HZNum[2].ToString();
                case '3':
                    return HZNum[3].ToString();
                case '4':
                    return HZNum[4].ToString();
                case '5':
                    return HZNum[5].ToString();
                case '6':
                    return HZNum[6].ToString();
                case '7':
                    return HZNum[7].ToString();
                case '8':
                    return HZNum[8].ToString();
                case '9':
                    return HZNum[9].ToString();
                default:
                    return "";
            }
        }
        #endregion

        #region BitTest32
        /// <summary>
        /// ??某位是否?真
        /// </summary>
        /// <param name="num"></param>
        /// <param name="bitpostion"></param>
        /// <returns></returns>
        private bool BitTest32(int num, int bitpostion)
        {

            if ((bitpostion > 31) || (bitpostion < 0))
                throw new Exception("Error Param: bitpostion[0-31]:" + bitpostion.ToString());

            int bit = 1 << bitpostion;

            if ((num & bit) == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        #endregion

        #region ConvertDayOfWeek
        /// <summary>
        /// ?星期幾?成?字表示
        /// </summary>
        /// <param name="dayOfWeek"></param>
        /// <returns></returns>
        private int ConvertDayOfWeek(DayOfWeek dayOfWeek)
        {
            switch (dayOfWeek)
            {
                case DayOfWeek.Sunday:
                    return 1;
                case DayOfWeek.Monday:
                    return 2;
                case DayOfWeek.Tuesday:
                    return 3;
                case DayOfWeek.Wednesday:
                    return 4;
                case DayOfWeek.Thursday:
                    return 5;
                case DayOfWeek.Friday:
                    return 6;
                case DayOfWeek.Saturday:
                    return 7;
                default:
                    return 0;
            }
        }
        #endregion

        #region CompareWeekDayHoliday
        /// <summary>
        /// 比??天是不是指定的第周幾
        /// </summary>
        /// <param name="date"></param>
        /// <param name="month"></param>
        /// <param name="week"></param>
        /// <param name="day"></param>
        /// <returns></returns>
        private bool CompareWeekDayHoliday(DateTime date, int month, int week, int day)
        {
            bool ret = false;

            if (date.Month == month) //月份相同
            {
                if (ConvertDayOfWeek(date.DayOfWeek) == day) //星期幾相同
                {
                    DateTime firstDay = new DateTime(date.Year, date.Month, 1);//生成?月第一天
                    int i = ConvertDayOfWeek(firstDay.DayOfWeek);
                    int firWeekDays = 7 - ConvertDayOfWeek(firstDay.DayOfWeek) + 1; //?算第一周剩餘天?

                    if (i > day)
                    {
                        if ((week - 1) * 7 + day + firWeekDays == date.Day)
                        {
                            ret = true;
                        }
                    }
                    else
                    {
                        if (day + firWeekDays + (week - 2) * 7 == date.Day)
                        {
                            ret = true;
                        }
                    }
                }
            }

            return ret;
        }
        #endregion
        #endregion

        #region  ?性

        #region ?日
        #region newCalendarHoliday
        /// <summary>
        /// ?算中????日
        /// </summary>
        public string newCalendarHoliday
        {
            get
            {
                string tempStr = "";
                if (this._cIsLeapMonth == false) //?月不?算?日
                {
                    foreach (LunarHolidayStruct lh in lHolidayInfo)
                    {
                        if ((lh.Month == this._cMonth) && (lh.Day == this._cDay))
                        {

                            tempStr = lh.HolidayName;
                            break;

                        }
                    }

                    //?除夕?行特??理
                    if (this._cMonth == 12)
                    {
                        int i = GetChineseMonthDays(this._cYear, 12); //?算?年??12月的?天?
                        if (this._cDay == i) //如果?最後一天
                        {
                            tempStr = "除夕";
                        }
                    }
                }
                return tempStr;
            }
        }
        #endregion

        #region WeekDayHoliday
        /// <summary>
        /// 按某月第幾周第幾日?算的?日
        /// </summary>
        public string WeekDayHoliday
        {
            get
            {
                string tempStr = "";
                foreach (WeekHolidayStruct wh in wHolidayInfo)
                {
                    if (CompareWeekDayHoliday(_date, wh.Month, wh.WeekAtMonth, wh.WeekDay))
                    {
                        tempStr = wh.HolidayName;
                        break;
                    }
                }
                return tempStr;
            }
        }
        #endregion

        #region DateHoliday
        /// <summary>
        /// 按公?日?算的?日
        /// </summary>
        public string DateHoliday
        {
            get
            {
                string tempStr = "";

                foreach (SolarHolidayStruct sh in sHolidayInfo)
                {
                    if ((sh.Month == _date.Month) && (sh.Day == _date.Day))
                    {
                        tempStr = sh.HolidayName;
                        break;
                    }
                }
                return tempStr;
            }
        }
        #endregion
        #endregion

        #region 公?日期
        #region Date
        /// <summary>
        /// 取??的公?日期
        /// </summary>
        public DateTime Date
        {
            get { return _date; }
            set { _date = value; }
        }
        #endregion

        #region WeekDay
        /// <summary>
        /// 取星期幾
        /// </summary>
        public DayOfWeek WeekDay
        {
            get { return _date.DayOfWeek; }
        }
        #endregion

        #region WeekDayStr
        /// <summary>
        /// 周幾的字元
        /// </summary>
        public string WeekDayStr
        {
            get
            {
                switch (_date.DayOfWeek)
                {
                    case DayOfWeek.Sunday:
                        return "星期日";
                    case DayOfWeek.Monday:
                        return "星期一";
                    case DayOfWeek.Tuesday:
                        return "星期二";
                    case DayOfWeek.Wednesday:
                        return "星期三";
                    case DayOfWeek.Thursday:
                        return "星期四";
                    case DayOfWeek.Friday:
                        return "星期五";
                    default:
                        return "星期六";
                }
            }
        }
        #endregion

        #region DateString
        /// <summary>
        /// 公?日期中文標記法 如一九九七年七月一日
        /// </summary>
        public string DateString
        {
            get
            {
                return "西元" + this._date.ToLongDateString();
            }
        }
        #endregion

        #region IsLeapYear
        /// <summary>
        /// ?前是否公??年
        /// </summary>
        public bool IsLeapYear
        {
            get
            {
                return DateTime.IsLeapYear(this._date.Year);
            }
        }
        #endregion

        #region ChineseConstellation
        /// <summary>
        /// 28星宿?算
        /// </summary>
        public string ChineseConstellation
        {
            get
            {
                int offset = 0;
                int modStarDay = 0;

                TimeSpan ts = this._date - ChineseConstellationReferDay;
                offset = ts.Days;
                modStarDay = offset % 28;
                return (modStarDay >= 0 ? _chineseConstellationName[modStarDay] : _chineseConstellationName[27 + modStarDay]);
            }
        }
        #endregion

        #region ChineseHour
        /// <summary>
        /// ?辰
        /// </summary>
        public string ChineseHour
        {
            get
            {
                return GetChineseHour(_datetime);
            }
        }
        #endregion

        #endregion

        #region ??日期
        #region IsChineseLeapMonth
        /// <summary>
        /// 是否?月
        /// </summary>
        public bool IsChineseLeapMonth
        {
            get { return this._cIsLeapMonth; }
        }
        #endregion

        #region IsChineseLeapYear
        /// <summary>
        /// ?年是否有?月
        /// </summary>
        public bool IsChineseLeapYear
        {
            get
            {
                return this._cIsLeapYear;
            }
        }
        #endregion

        #region ChineseDay
        /// <summary>
        /// ??日
        /// </summary>
        public int ChineseDay
        {
            get { return this._cDay; }
        }
        #endregion

        #region ChineseDayString
        /// <summary>
        /// ??日中文表示
        /// </summary>
        public string ChineseDayString
        {
            get
            {
                switch (this._cDay)
                {
                    case 0:
                        return "";
                    case 10:
                        return "初十";
                    case 20:
                        return "二十";
                    case 30:
                        return "三十";
                    default:
                        return nStr2[(int)(_cDay / 10)].ToString() + nStr1[_cDay % 10].ToString();

                }
            }
        }
        #endregion

        #region ChineseMonth
        /// <summary>
        /// ??的月份
        /// </summary>
        public int ChineseMonth
        {
            get { return this._cMonth; }
        }
        #endregion

        #region ChineseMonthString
        /// <summary>
        /// ??月份字串
        /// </summary>
        public string ChineseMonthString
        {
            get
            {
                return _monthString[this._cMonth];
            }
        }
   
        #endregion

        #region ChineseYear
        /// <summary>
        /// 取??年份
        /// </summary>
        public int ChineseYear
        {
            get { return this._cYear; }
        }
        #endregion

        #region ChineseYearString
        /// <summary>
        /// 取??年字串如，一九九七年
        /// </summary>
        public string ChineseYearString
        {
            get
            {
                string tempStr = "";
                string num = this._cYear.ToString();
                for (int i = 0; i < 4; i++)
                {
                    tempStr += ConvertNumToChineseNum(num[i]);
                }
                return tempStr + "年";
            }
        }
        #endregion

        #region ChineseDateString
        /// <summary>
        /// 取??日期標記法：??一九九七年正月初五
        /// </summary>
        public string ChineseDateString
        {
            get
            {
                if (this._cIsLeapMonth == true)
                {
                    return "農" + ChineseYearString + "月" + ChineseMonthString + ChineseDayString;
                }
                else
                {
                    return "農" + ChineseYearString + ChineseMonthString + ChineseDayString;
                }
            }
        }
        public string TaiwanDateString
        {
            get
            {
                if (this._cIsLeapMonth == true)
                {
                    int TaiwanYear = ChineseYear - 1911;
                    return TaiwanYear.ToString("000") + ChineseMonth.ToString("00") + ChineseDay.ToString("00");
                }
                else
                {
                    int TaiwanYear = ChineseYear - 1911;
                    return TaiwanYear.ToString("000") + ChineseMonth.ToString("00") + ChineseDay.ToString("00");
                }
            }
        }
        #endregion

        #region ChineseTwentyFourDay
        /// <summary>
        /// 定氣法?算二十四?氣,二十四?氣是按地球公???算的，並非是???算的
        /// </summary>
        /// <remarks>
        /// ?氣的定法有?種。古代?法採用的??"恒氣"，即按??把一年等分?24份，
        /// 每一?氣平均得15天有餘，所以又?"平氣"。?代??採用的??"定氣"，即
        /// 按地球在?道上的位置??准，一周360°，??氣之?相隔15°。由於冬至?地
        /// 球位於近日?附近，??速度?快，因而太?在?道上移?15°的??不到15天。
        /// 夏至前後的情?正好相反，太?在?道上移??慢，一??氣?16天之多。採用
        /// 定氣?可以保?春、秋?分必然在?夜平分的那?天。
        /// </remarks>
        //public string ChineseTwentyFourDay
        //{
        //    get
        //    {
        //        DateTime baseDateAndTime = new DateTime(1900, 1, 6, 2, 5, 0); //#1/6/1900 2:05:00 AM#
        //        DateTime newDate;
        //        double num;
        //        int y;
        //        string tempStr = "";

        //        y = this._date.Year;

        //        for (int i = 24; i >= 1; i--)
        //        {
        //            num = 525948.76 * (y - 1900) + sTermInfo[i - 1];

        //            newDate = baseDateAndTime.AddMinutes(num);//按分??算

        //            if (newDate.DayOfYear < _date.DayOfYear)
        //            {
        //                tempStr = string.Format("{0}[{1}]", SolarTerm[i - 1], newDate.ToString("yyyy-MM-dd"));
        //                break;
        //            }
        //        }

        //        if (tempStr == "")
        //        {
        //            y = y - 1;

        //            num = 525948.76 * (y - 1900) + sTermInfo[24 - 1];

        //            newDate = baseDateAndTime.AddMinutes(num);//按分??算
        //            tempStr = string.Format("{0}[{1}]", SolarTerm[24 - 1], newDate.ToString("yyyy-MM-dd"));

        //        }
        //        int seg24 = 0;
        //                switch (tempStr)
        //                {
        //                    case "小寒": seg24 = 23; break;
        //                    case "大寒": seg24 = 24; break; 
        //                    case "立春": seg24 = 1 ; break;  
        //                    case "雨水": seg24 = 2; break;
        //                    case "驚蟄": seg24 = 3; break;
        //                    case "春分": seg24 = 4; break;
        //                    case "清明": seg24 = 5; break;
        //                    case "穀雨": seg24 = 6; break;
        //                    case "立夏": seg24 = 7; break;
        //                    case "小滿": seg24 = 8; break;
        //                    case "芒種": seg24 = 9; break;
        //                    case "夏至": seg24 = 10; break;
        //                    case "小暑": seg24 = 11; break;
        //                    case "大暑": seg24 = 12; break;
        //                    case "立秋": seg24 = 13; break;
        //                    case "處暑": seg24 = 14; break;
        //                    case "白露": seg24 = 15; break;
        //                    case "秋分": seg24 = 16; break;
        //                    case "寒露": seg24 = 17; break;
        //                    case "霜降": seg24 = 18; break;
        //                    case "立冬": seg24 = 19; break;
        //                    case "小雪": seg24 = 20; break;
        //                    case "大雪": seg24 = 21; break;
        //                    case "冬至": seg24 = 22; break;                          
        //                }
        //                int mh = seg24 / 2;//结果为商1
        //                int d = seg24 % 2;//结果为余数
        //                if (d>0)
        //                {
        //                    mh = mh + 1;
        //                }
        //                //tempStr = string.Format("{0}[{1}]", mh.ToString(), newDate.ToString("yyyy-MM-dd"));
        //                tempStr = mh.ToString();
        //                //tempStr = string.Format("{0}[{1}]", SolarTerm[i - 1], newDate.ToString("yyyy-MM-dd"));

        //        return tempStr;
        //    }
        //}

        public string ChineseTwentyFourPrev2
        {
            get
            {
                DateTime baseDateAndTime = new DateTime(1900, 1, 6, 2, 5, 0); //#1/6/1900 2:05:00 AM#
                DateTime newDate;
                double num;
                int y;
                int _prev = 0;
                string tempStr = "";

                y = this._date.Year;

                for (int i = 24; i >= 1; i--)
                {
                    num = 525948.76 * (y - 1900) + sTermInfo[i - 1];

                    newDate = baseDateAndTime.AddMinutes(num);//按分??算

                    if (newDate.DayOfYear < _date.DayOfYear)
                    {
                        _prev = _prev + 1;
                        tempStr = string.Format("{0}[{1}]", SolarTerm[i - 1], newDate.ToString("yyyy-MM-dd"));
                        if (_prev == 2)
                        {
                            break;
                        }
                    }
                }

                if (tempStr == "")
                {
                    y = y - 1;

                    num = 525948.76 * (y - 1900) + sTermInfo[24 - 1];

                    newDate = baseDateAndTime.AddMinutes(num);//按分??算
                    tempStr = string.Format("{0}[{1}]", SolarTerm[24 - 1], newDate.ToString("yyyy-MM-dd"));

                }


                return tempStr;
            }

        }

        //?前日期前一?最近?氣
        //public string ChineseTwentyFourPrevDay
        //{
        //    get
        //    {
        //        DateTime baseDateAndTime = new DateTime(1900, 1, 6, 2, 5, 0); //#1/6/1900 2:05:00 AM#
        //        DateTime newDate;
        //        double num;
        //        int y;
        //        string tempStr = "";

        //        y = this._date.Year;

        //        for (int i = 24; i >= 1; i--)
        //        {
        //            num = 525948.76 * (y - 1900) + sTermInfo[i - 1];

        //            newDate = baseDateAndTime.AddMinutes(num);//按分??算

        //            if (newDate.DayOfYear < _date.DayOfYear)
        //            {
        //                tempStr = string.Format("{0}[{1}]", SolarTerm[i - 1], newDate.ToString("yyyy-MM-dd"));
        //                break;
        //            }
        //        }

        //        if (tempStr=="")
        //        {
        //            y = y-1;

        //                num = 525948.76 * (y - 1900) + sTermInfo[24 - 1];

        //                newDate = baseDateAndTime.AddMinutes(num);//按分??算
        //               tempStr = string.Format("{0}[{1}]", SolarTerm[24 - 1], newDate.ToString("yyyy-MM-dd"));

        //         }


        //        return tempStr;
        //    }

        //}

        //?前日期前一?最近?氣

        public int ChineseTwentyFourP2
        {
            get
            {
                DateTime baseDateAndTime = new DateTime(1900, 1, 6, 2, 5, 0); //#1/6/1900 2:05:00 AM#
                DateTime newDate;
                double num;
                int y;
                int tempStr = 0;
                int _Prev = 0;
                y = this._date.Year;

                for (int i = 24; i >= 1; i--)
                {
                    num = 525948.76 * (y - 1900) + sTermInfo[i - 1];

                    newDate = baseDateAndTime.AddMinutes(num);//按分??算

                    if (newDate.DayOfYear < _date.DayOfYear)
                    {
                        _Prev = _Prev + 1;
                        tempStr = newDate.DayOfYear;
                        if (_Prev == 2)
                        {
                            break;
                        }
                    }
                }

                if (tempStr == 0)
                {
                    y = y - 1;

                    num = 525948.76 * (y - 1900) + sTermInfo[24 - 1];

                    newDate = baseDateAndTime.AddMinutes(num);//按分??算
                    tempStr = newDate.DayOfYear;

                }


                return tempStr;
            }

        }

        //public int ChineseTwentyFourPrev
        //{
        //    get
        //    {
        //        DateTime baseDateAndTime = new DateTime(1900, 1, 6, 2, 5, 0); //#1/6/1900 2:05:00 AM#
        //        DateTime newDate;
        //        double num;
        //        int y;
        //        int tempStr = 0;

        //        y = this._date.Year;

        //        for (int i = 24; i >= 1; i--)
        //        {
        //            num = 525948.76 * (y - 1900) + sTermInfo[i - 1];

        //            newDate = baseDateAndTime.AddMinutes(num);//按分??算

        //            if (newDate.DayOfYear < _date.DayOfYear)
        //            {
        //                tempStr = newDate.DayOfYear;
        //                break;
        //            }
        //        }

        //        if (tempStr == 0)
        //        {
        //            y = y - 1;

        //            num = 525948.76 * (y - 1900) + sTermInfo[24 - 1];

        //            newDate = baseDateAndTime.AddMinutes(num);//按分??算
        //            tempStr = newDate.DayOfYear;

        //        }


        //        return tempStr;
        //    }

        //}

        public int ChineseTwentyFourN2
        {
            get
            {
                DateTime baseDateAndTime = new DateTime(1900, 1, 6, 2, 5, 0); //#1/6/1900 2:05:00 AM#
                DateTime newDate;
                double num;
                int y;
                int _Next = 0;
                int tempStr = 0;

                y = this._date.Year;

                for (int i = 1; i <= 24; i++)
                {
                    num = 525948.76 * (y - 1900) + sTermInfo[i - 1];

                    newDate = baseDateAndTime.AddMinutes(num);//按分??算

                    if (newDate.DayOfYear > _date.DayOfYear)
                    {
                        _Next = _Next+1;
                        tempStr = newDate.DayOfYear;
                        if (_Next == 2)
                        {
                            break;
                        }
                    }
                }
                if (tempStr == 0)
                {
                    y = y + 1;

                    num = 525948.76 * (y - 1900) + sTermInfo[1 - 1];

                    newDate = baseDateAndTime.AddMinutes(num);//按分??算
                    tempStr = newDate.DayOfYear;

                }
                return tempStr;
            }

        }

        //?前日期後一?最近?氣
        //public int ChineseTwentyFourNext
        //{
        //    get
        //    {
        //        DateTime baseDateAndTime = new DateTime(1900, 1, 6, 2, 5, 0); //#1/6/1900 2:05:00 AM#
        //        DateTime newDate;
        //        double num;
        //        int y;
        //        int tempStr = 0;

        //        y = this._date.Year;

        //        for (int i = 1; i <= 24; i++)
        //        {
        //            num = 525948.76 * (y - 1900) + sTermInfo[i - 1];

        //            newDate = baseDateAndTime.AddMinutes(num);//按分??算

        //            if (newDate.DayOfYear > _date.DayOfYear)
        //            {
        //                tempStr = newDate.DayOfYear;
        //                break;
        //            }
        //        }
        //        if (tempStr == 0)
        //        {
        //            y = y + 1;

        //            num = 525948.76 * (y - 1900) + sTermInfo[1 - 1];

        //            newDate = baseDateAndTime.AddMinutes(num);//按分??算
        //            tempStr = newDate.DayOfYear;

        //        }
        //        return tempStr;
        //    }

        //}
        /// <summary>
        /// 取二十四節氣
        /// </summary>
        /// <param name="year">年</param>
        /// <param name="index">索引</param>
        /// <returns></returns>
        private static DateTime ChineseTwentyFourDay(int year, int index)
        {
            int[] SolarData = { 7, 21, 5, 20, 4, 20, 5, 21, 6, 22, 7, 23, 7, 23, 8, 23, 8, 23, 8, 24, 8, 23, 7, 22, 6, 22, 6, 21, 5, 21, 5, 20, 4, 20, 3, 20, 2, 19, 2, 18, 1, 20, 1, 20, 0, 20 };
            double[] SolarTerms = { 20.12, 2.18, 18.73, 2.83, 19.46, 6.11, 20.12, 6.78, 20.84, 7.39, 23.08, 7.54, 23.36, 9.47, 25.13, 9.77, 24.58, 10.15, 26.43, 10.96, 26.54, 12.06, 27.25, 12.98, 28.05, 14.5, 28.59, 15.63, 29.45, 16.51, 29.58, 17.58, 29.46, 18.73, 29.58, 20.12, 29.3, 20.83, 28.57, 21.8, 27.4, 22.33, 26.1, 22.84, 25.13, 23.04, 25.04, 23.08 };

            double c = 0.2422 * (year - GanZhiStartYear) - Math.Floor((double)((year - GanZhiStartYear) / 4));
            DateTime date = new DateTime(year, 1, 1).AddDays(SolarData[index * 2] + SolarTerms[index * 2] + (year == GanZhiStartYear ? SolarData[index * 2 + 1] + SolarTerms[index * 2 + 1] : 0) + c).AddDays(-1);
            return date.AddDays(1);
        }

        /// <summary>
        /// 取前一節氣
        /// </summary>
        /// <returns></returns>
        public DateTime ChineseTwentyFourPrevDay()
        {
            DateTime date = _date.AddDays(-1);
            int index = Array.IndexOf(SolarTerm, SolarTermString);
            if (index == -1)
            {
                for (int i = 0; i < SolarTerm.Length; i++)
                {
                    DateTime st = ChineseTwentyFourDay(_date.Year, i);
                    if (st.Date == _date.Date)
                    {
                        index = i;
                        break;
                    }
                }
            }

            if (index == -1)
            {
                return ChineseTwentyFourDay(date.Year, 23);
            }
            if (index == 0)
            {
                return ChineseTwentyFourDay(date.Year - 1, 23);
            }
            return ChineseTwentyFourDay(date.Year, index - 1);
        }

        /// <summary>
        /// 取下一節氣
        /// </summary>
        /// <returns></returns>
        public DateTime ChineseTwentyFourNext()
        {
            DateTime date = _date.AddDays(1);
            int index = Array.IndexOf(SolarTerm, SolarTermString);
            if (index == -1)
            {
                for (int i = 0; i < SolarTerm.Length; i++)
                {
                    DateTime st = ChineseTwentyFourDay(_date.Year, i);
                    if (st.Date == _date.Date)
                    {
                        index = i;
                        break;
                    }
                }
            }

            if (index == -1)
            {
                return ChineseTwentyFourDay(date.Year, 0);
            }
            if (index == 23)
            {
                return ChineseTwentyFourDay(date.Year + 1, 0);
            }
            return ChineseTwentyFourDay(date.Year, index + 1);
        }
        public string ChineseTwentyFourNext2
        {
            get
            {
                DateTime baseDateAndTime = new DateTime(1900, 1, 6, 2, 5, 0); //#1/6/1900 2:05:00 AM#
                DateTime newDate;
                double num;
                int y;
                int _Next = 0;
                string tempStr = "";

                y = this._date.Year;

                for (int i = 1; i <= 24; i++)
                {
                    num = 525948.76 * (y - 1900) + sTermInfo[i - 1];

                    newDate = baseDateAndTime.AddMinutes(num);//按分??算

                    if (newDate.DayOfYear > _date.DayOfYear)
                    {
                        _Next = _Next + 1;
                        tempStr = string.Format("{0}[{1}]", SolarTerm[i - 1], newDate.ToString("yyyy-MM-dd"));
                        if (_Next == 2)
                        {
                            break;
                        }
                    }
                }

                if (tempStr == "")
                {
                    y = y + 1;

                    num = 525948.76 * (y - 1900) + sTermInfo[1 - 1];

                    newDate = baseDateAndTime.AddMinutes(num);//按分??算
                    tempStr = string.Format("{0}[{1}]", SolarTerm[1 - 1], newDate.ToString("yyyy-MM-dd"));

                }
                return tempStr;
            }

        }

        //?前日期後一?最近?氣
        public string ChineseTwentyFourNextDay
        {
            get
            {
                DateTime baseDateAndTime = new DateTime(1900, 1, 6, 2, 5, 0); //#1/6/1900 2:05:00 AM#
                DateTime newDate;
                double num;
                int y;
                string tempStr = "";

                y = this._date.Year;

                for (int i = 1; i <= 24; i++)
                {
                    num = 525948.76 * (y - 1900) + sTermInfo[i - 1];

                    newDate = baseDateAndTime.AddMinutes(num);//按分??算

                    if (newDate.DayOfYear > _date.DayOfYear)
                    {
                        tempStr = string.Format("{0}[{1}]", SolarTerm[i - 1], newDate.ToString("yyyy-MM-dd"));
                        break;
                    }
                }

                if (tempStr == "")
                {
                    y = y + 1;

                    num = 525948.76 * (y - 1900) + sTermInfo[1 - 1];

                    newDate = baseDateAndTime.AddMinutes(num);//按分??算
                    tempStr = string.Format("{0}[{1}]", SolarTerm[1 - 1], newDate.ToString("yyyy-MM-dd"));

                }
                return tempStr;
            }

        }


        public string[] ChineseTwentyFour
        {
            get
            {
                DateTime baseDateAndTime = new DateTime(1900, 1, 6, 2, 5, 0); //#1/6/1900 2:05:00 AM#
                DateTime newDate;
                int seg24 = 0;
                string[] g24 = new string[25]; //節氣
                double num;
                int y;
                y = this._date.Year;

                for (int i = 24; i >= 1; i--)
                {
                    num = 525948.76 * (y - 1900) + sTermInfo[i - 1];

                    newDate = baseDateAndTime.AddMinutes(num);//按分??算

                    switch (SolarTerm[i - 1].ToString().Trim())
                    {
                        case "小寒": seg24 = 23; break;
                        case "大寒": seg24 = 24; break;
                        case "立春": seg24 = 1; break;
                        case "雨水": seg24 = 2; break;
                        case "驚蟄": seg24 = 3; break;
                        case "春分": seg24 = 4; break;
                        case "清明": seg24 = 5; break;
                        case "穀雨": seg24 = 6; break;
                        case "立夏": seg24 = 7; break;
                        case "小滿": seg24 = 8; break;
                        case "芒種": seg24 = 9; break;
                        case "夏至": seg24 = 10; break;
                        case "小暑": seg24 = 11; break;
                        case "大暑": seg24 = 12; break;
                        case "立秋": seg24 = 13; break;
                        case "處暑": seg24 = 14; break;
                        case "白露": seg24 = 15; break;
                        case "秋分": seg24 = 16; break;
                        case "寒露": seg24 = 17; break;
                        case "霜降": seg24 = 18; break;
                        case "立冬": seg24 = 19; break;
                        case "小雪": seg24 = 20; break;
                        case "大雪": seg24 = 21; break;
                        case "冬至": seg24 = 22; break;
                    }

                    g24[seg24] = newDate.ToString(("yyyy/MM/dd HH:mm:ss"));
                    //}
                }

  
                return g24;
            }
        }

        #endregion
        #endregion


        //public DateTime GetNextLunarNewYearDate()
        //{
        //    ushort iYear, iMonth, iDay;

        //    TimeSpan ts;
        //    iYear = (ushort)(m_Date.AddYears(1).Year); //明年

        //    if ((iYear < START_YEAR) || (iYear > END_YEAR)) { return DateTime.MinValue; };

        //    ts = m_Date.AddYears(1) - (new DateTime(START_YEAR, 1, 1));
        //    l_CalcLunarDate(out iYear, out iMonth, out iDay, (uint)(ts.Days));

        //    // 今年過年至今的農曆天數
        //    uint days = 0;

        //    for (ushort i = 1; i < iMonth; i++)
        //    {
        //        days += LunarMonthDays(iYear, i);
        //    }

        //    days += iDay - (uint)1;

        //    return m_Date.AddYears(1).AddDays(-days);
        //}

        #region 星座
        #region Constellation
        /// <summary>
        /// ?算指定日期的星座序? 
        /// </summary>
        /// <returns></returns>
        public string Constellation
        {
            get
            {
                int index = 0;
                int y, m, d;
                y = _date.Year;
                m = _date.Month;
                d = _date.Day;
                y = m * 100 + d;

                if (((y >= 321) && (y <= 419))) { index = 0; }
                else if ((y >= 420) && (y <= 520)) { index = 1; }
                else if ((y >= 521) && (y <= 620)) { index = 2; }
                else if ((y >= 621) && (y <= 722)) { index = 3; }
                else if ((y >= 723) && (y <= 822)) { index = 4; }
                else if ((y >= 823) && (y <= 922)) { index = 5; }
                else if ((y >= 923) && (y <= 1022)) { index = 6; }
                else if ((y >= 1023) && (y <= 1121)) { index = 7; }
                else if ((y >= 1122) && (y <= 1221)) { index = 8; }
                else if ((y >= 1222) || (y <= 119)) { index = 9; }
                else if ((y >= 120) && (y <= 218)) { index = 10; }
                else if ((y >= 219) && (y <= 320)) { index = 11; }
                else { index = 0; }

                return _constellationName[index];
            }
        }
        #endregion
        #endregion

        #region ?相
        #region Animal
        /// <summary>
        /// ?算?相的索引，注意?然?相是以??年???的，但是目前在??使用中是按公???算的
        /// 鼠年?1,其它?推
        /// </summary>
        public int Animal
        {
            get
            {
                int offset = _date.Year - AnimalStartYear;
                return (offset % 12) + 1;
            }
        }
        #endregion

        #region AnimalString
        /// <summary>
        /// 取?相字串
        /// </summary>
        public string AnimalString
        {
            get
            {
                int offset = _date.Year - AnimalStartYear; //???算
                //int offset = this._cYear - AnimalStartYear;　???算
                return animalStr[offset % 12].ToString();
            }
        }
        #endregion
        #endregion

        #region 天干地支
        #region GanZhiYearString
        /// <summary>
        /// 取??年的干支標記法如 乙丑年
        /// </summary>
        public string GanZhiYearString
        {
            get
            {
                string tempStr;
                int i = (this._cYear - GanZhiStartYear) % 60; //?算干支
                tempStr = ganStr[i % 10].ToString() + zhiStr[i % 12].ToString() + "年";
                return tempStr;
            }
        }
        #endregion

        public string GanZhiNewMonth(int _cMonth)
        {

                //每?月的地支?是固定的,而且?是?寅月?始
                int zhiIndex;
                string zhi;
                if (_cMonth > 10)
                {
                    zhiIndex = _cMonth - 10;
                }
                else
                {
                    zhiIndex = _cMonth + 2;
                }
                //zhi = zhiStr[zhiIndex - 1].ToString();
                zhi = (zhiIndex).ToString("00");
                //根據?年的干支年的幹??算月幹的第一?
                int ganIndex = 1;
                string gan;
                int i = (this._cYear - GanZhiStartYear) % 60; //?算干支
                switch (i % 10)
                {
                    #region ...
                    case 0: //甲
                        ganIndex = 3;
                        break;
                    case 1: //乙
                        ganIndex = 5;
                        break;
                    case 2: //丙
                        ganIndex = 7;
                        break;
                    case 3: //丁
                        ganIndex = 9;
                        break;
                    case 4: //戊
                        ganIndex = 1;
                        break;
                    case 5: //己
                        ganIndex = 3;
                        break;
                    case 6: //庚
                        ganIndex = 5;
                        break;
                    case 7: //辛
                        ganIndex = 7;
                        break;
                    case 8: //壬
                        ganIndex = 9;
                        break;
                    case 9: //癸
                        ganIndex = 1;
                        break;
                        #endregion
                }
                //gan = ganStr[(ganIndex + this._cMonth - 2) % 10].ToString();
                int xx = (ganIndex + _cMonth - 2) % 10;
                gan = (xx + 1).ToString("00");
                return gan + zhi + "月";
 
        }
        public string GanZhiMonth
        {
            get
            {
                //每?月的地支?是固定的,而且?是?寅月?始
                int zhiIndex;
                string zhi;
                if (this._cMonth > 10)
                {
                    zhiIndex = this._cMonth - 10;
                }
                else
                {
                    zhiIndex = this._cMonth + 2;
                }
                //zhi = zhiStr[zhiIndex - 1].ToString();
                zhi = (zhiIndex).ToString("00");
                //根據?年的干支年的幹??算月幹的第一?
                int ganIndex = 1;
                string gan;
                int i = (this._cYear - GanZhiStartYear) % 60; //?算干支
                switch (i % 10)
                {
                    #region ...
                    case 0: //甲
                        ganIndex = 3;
                        break;
                    case 1: //乙
                        ganIndex = 5;
                        break;
                    case 2: //丙
                        ganIndex = 7;
                        break;
                    case 3: //丁
                        ganIndex = 9;
                        break;
                    case 4: //戊
                        ganIndex = 1;
                        break;
                    case 5: //己
                        ganIndex = 3;
                        break;
                    case 6: //庚
                        ganIndex = 5;
                        break;
                    case 7: //辛
                        ganIndex = 7;
                        break;
                    case 8: //壬
                        ganIndex = 9;
                        break;
                    case 9: //癸
                        ganIndex = 1;
                        break;
                        #endregion
                }
                //gan = ganStr[(ganIndex + this._cMonth - 2) % 10].ToString();
                int xx = (ganIndex + this._cMonth - 2) % 10;
                gan = (xx+1).ToString("00");
                return gan + zhi + "月";
            }
        }
        #region GanZhiMonthString
        /// <summary>
        /// 取干支的月表示字串，注意??的?月不?干支
        /// </summary>
        public string GanZhiMonthString
        {
            get
            {
                //每?月的地支?是固定的,而且?是?寅月?始
                int zhiIndex;
                string zhi;
                if (this._cMonth > 10)
                {
                    zhiIndex = this._cMonth - 10;
                }
                else
                {
                    zhiIndex = this._cMonth + 2;
                }
                zhi = zhiStr[zhiIndex - 1].ToString();

                //根據?年的干支年的幹??算月幹的第一?
                int ganIndex = 1;
                string gan;
                int i = (this._cYear - GanZhiStartYear) % 60; //?算干支
                switch (i % 10)
                {
                    #region ...
                    case 0: //甲
                        ganIndex = 3;
                        break;
                    case 1: //乙
                        ganIndex = 5;
                        break;
                    case 2: //丙
                        ganIndex = 7;
                        break;
                    case 3: //丁
                        ganIndex = 9;
                        break;
                    case 4: //戊
                        ganIndex = 1;
                        break;
                    case 5: //己
                        ganIndex = 3;
                        break;
                    case 6: //庚
                        ganIndex = 5;
                        break;
                    case 7: //辛
                        ganIndex = 7;
                        break;
                    case 8: //壬
                        ganIndex = 9;
                        break;
                    case 9: //癸
                        ganIndex = 1;
                        break;
                    #endregion
                }
                gan = ganStr[(ganIndex + this._cMonth - 2) % 10].ToString();

                return gan + zhi + "月";
            }
        }
  
        #endregion

        #region GanZhiDayString
        /// <summary>
        /// 取干支日標記法
        /// </summary>
        public string GanZhiDayString
        {
            get
            {
                int i, offset;
                TimeSpan ts = this._date - GanZhiStartDay;
                offset = ts.Days;
                i = offset % 60;
                return ganStr[i % 10].ToString() + zhiStr[i % 12].ToString() + "日";
            }
        }
        #endregion

        #region GanZhiDateString
        /// <summary>
        /// 取?前日期的干支標記法如 甲子年乙丑月丙庚日
        /// </summary>
        public string GanZhiDateString
        {
            get
            {
                return GanZhiYearString + GanZhiMonthString + GanZhiDayString;
            }
        }
        public string GanZhiYYString
        {
            get
            {
                return GanZhiYearString;
            }
        }

        public string GanZhiMMString
        {
            get
            {
                return   GanZhiMonthString  ;
            }
        }
        public string GanZhiNewMMString(int _cMonth)
        {
                //每?月的地支?是固定的,而且?是?寅月?始
                int zhiIndex;
                string zhi;
                if ( _cMonth > 10)
                {
                    zhiIndex =  _cMonth - 10;
                }
                else
                {
                    zhiIndex =  _cMonth + 2;
                }
                    zhi = zhiStr[zhiIndex - 1].ToString();

                    //根據?年的干支年的幹??算月幹的第一?
                    int ganIndex = 1;
                    string gan;
                    int i = (this._cYear - GanZhiStartYear) % 60; //?算干支
                                switch (i % 10)
                                {
                                    #region ...
                                    case 0: //甲
                                        ganIndex = 3;
                                        break;
                                    case 1: //乙
                                        ganIndex = 5;
                                        break;
                                    case 2: //丙
                                        ganIndex = 7;
                                        break;
                                    case 3: //丁
                                        ganIndex = 9;
                                        break;
                                    case 4: //戊
                                        ganIndex = 1;
                                        break;
                                    case 5: //己
                                        ganIndex = 3;
                                        break;
                                    case 6: //庚
                                        ganIndex = 5;
                                        break;
                                    case 7: //辛
                                        ganIndex = 7;
                                        break;
                                    case 8: //壬
                                        ganIndex = 9;
                                        break;
                                    case 9: //癸
                                        ganIndex = 1;
                                        break;
                                    #endregion
                                }
                gan = ganStr[(ganIndex + _cMonth - 2) % 10].ToString();

                                return gan + zhi + "月";
        }
        public string GanZhiDDString
        {
            get
            {
                return  GanZhiDayString;
            }
        }
        #endregion
        #endregion
        #endregion

        #region 方法
        #region NextDay
        /// <summary>
        /// 取下一天
        /// </summary>
        /// <returns></returns>
        public EcanChineseCalendar NextDay()
        {
            DateTime nextDay = _date.AddDays(1);
            return new EcanChineseCalendar(nextDay);
        }
        #endregion

        #region PervDay
        /// <summary>
        /// 取前一天
        /// </summary>
        /// <returns></returns>
        public EcanChineseCalendar PervDay()
        {
            DateTime pervDay = _date.AddDays(-1);
            return new EcanChineseCalendar(pervDay);
        }
        #endregion
        #endregion
    }
}

