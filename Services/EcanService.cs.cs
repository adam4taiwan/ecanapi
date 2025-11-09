using Ecan;
using Ecanapi.Models.Ecanapi.Models;
using Ecanapi.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;

namespace Ecanapi.Services
{
    public class EcanService
    {
        // 使用用戶提供的 EcanChineseCalendar 進行農曆和八字計算
        public BaziData CalculateBazi(int year, int month, int day, int hour, bool isSolar = true)
        {
            // 假設輸入為公曆日期，創建 EcanChineseCalendar 實例
            System.DateTime dt2 = new System.DateTime(year, month, day, hour, 0, 0);
            var dt = new DateTime(year, month, day, hour, 0, 0);
            var chineseCalendar = new EcanChineseCalendar(dt);
            // 24 節氣
            string[] g24 = chineseCalendar.ChineseTwentyFour;
            var baziData = new BaziData();

            // 解析年干支 (e.g., "甲子年" -> YearGan = "甲", YearZhi = "子")
            string ganZhiYear = chineseCalendar.GanZhiYYString; // 注意：用戶類中是 GanZhiYearString
            baziData.YearGan = ganZhiYear.Substring(0, 1);
            baziData.YearZhi = ganZhiYear.Substring(1, 1);

            // 解析月干支
            string ganZhiMonth = chineseCalendar.GanZhiMMString;
            baziData.MonthGan = ganZhiMonth.Substring(0, 1);
            baziData.MonthZhi = ganZhiMonth.Substring(1, 1);

            // 解析日干支
            string ganZhiDay = chineseCalendar.GanZhiDDString.TrimEnd('日');
            baziData.DayGan = ganZhiDay.Substring(0, 1);
            baziData.DayZhi = ganZhiDay.Substring(1, 1);

            // 解析時干支 (ChineseHour 返回如 "甲子")
            string chineseHour = chineseCalendar.ChineseHour;
            baziData.HourGan = chineseHour.Substring(0, 1);
            baziData.HourZhi = chineseHour.Substring(1, 1);

            // 從 SiZiConstants 獲取摘要
            string key = $"{baziData.DayGan}日{baziData.HourGan}{baziData.HourZhi}";
            baziData.Summary = SiZiConstants.Summarys.ContainsKey(key) ? SiZiConstants.Summarys[key] : "無匹配摘要";

            // 實現 common.py 的 check_gan 邏輯
            string[] gans = { baziData.YearGan, baziData.MonthGan, baziData.DayGan, baziData.HourGan };
            string[] zhis = { baziData.YearZhi, baziData.MonthZhi, baziData.DayZhi, baziData.HourZhi };
            var genResult = GetGen(baziData.DayGan, zhis);

            // 實現 yinyang 邏輯（從 common.py）
            string yinYang = YinYang(baziData.DayZhi);

            // 實現 luohou.py 的 get_hou 邏輯
            // 需要夏至和冬至，使用 EcanChineseCalendar 的節氣計算
            // 假設使用 SolarTermString 或 GetLunarHolDay 來獲取夏至/冬至
            DateTime xiazhi = chineseCalendar.ChineseTwentyFourNext(); // 示例，需調整為實際夏至日期
            DateTime dongzhi = chineseCalendar.ChineseTwentyFourNext(); // 示例，需調整
            // string luohou = GetHou(day, xiazhi, dongzhi);

            return baziData;
        }

        // 對應 common.py 的 check_gan
        private string CheckGan(string gan, string[] gans)
        {
            var counts = new Dictionary<string, int>();
            foreach (var g in GanZhiConstants.Gan)
                counts[g] = gans.Count(x => x == g);

            string result = "";
            if (counts[gan] >= 2)
                result += $"{gan}多，";
            else if (counts[gan] == 0)
                result += $"{gan}無，";

            // 其他十神邏輯（根據 Python common.py）
            // 例如：比肩、劫財、正印等
            return result.TrimEnd('，');
        }

        // 對應 common.py 的 get_gen
        private string GetGen(string dayGan, string[] zhis)
        {
            string result = "";
            foreach (var zhi in zhis)
            {
                if (GanZhiConstants.Zhi5List.ContainsKey(zhi))
                {
                    var gansInZhi = GanZhiConstants.Zhi5List[zhi];
                    foreach (var gan in gansInZhi)
                    {
                        // 根據日干和地支藏干計算十神
                        result += $"{gan}在{zhi}，";
                    }
                }
            }
            return result.TrimEnd('，');
        }

        // 對應 common.py 的 yinyang
        private string YinYang(string zhi)
        {
            string[] yin = { "丑", "卯", "巳", "未", "酉", "亥" };
            return yin.Contains(zhi) ? "陰" : "陽";
        }

        // 對應 luohou.py 的 get_hou（使用 EcanChineseCalendar 的節氣計算）
        private string GetHou(int day, DateTime xiazhi, DateTime dongzhi)
        {
            // 假設需要夏至和冬至日期來計算羅喉
            // 實際邏輯需根據 luohou.py 補充，使用 EcanChineseCalendar 的 ChineseTwentyFourDay 等方法獲取節氣
            return "羅喉計算待實現";
        }

        // 對應 shengxiao.py 的生肖計算
        private string GetShengXiao(int year)
        {
            string[] shengXiao = { "鼠", "牛", "虎", "兔", "龍", "蛇", "馬", "羊", "猴", "雞", "狗", "豬" };
            return shengXiao[(year - 4) % 12];
        }

        // 其他方法（如 convert.py 的轉換邏輯）待補充
    }
}