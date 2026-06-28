namespace Ecanapi.Models;

/// <summary>
/// 定義 API 回傳給客戶端的 JSON 資料格式。
/// 使用 record 型別可以讓程式碼更簡潔且具備不可變性。
/// </summary>
public record CalendarResponse(
    int Year,
    int SolarMonth,
    int SolarDay,
    string? YearGanzhi,
    string? MonthGanzhi,
    string? DayGanzhi,
    string? DayTianGan,
    string? SolarTerm,
    string? LunarMonth,
    string? LunarDay,
    string? WeekDay,
    string? Season,
    string? CurrentSolarTerm,      // 當天所屬節氣名稱（當天若有節氣則同 SolarTerm，否則為前一個節氣）
    string? CurrentSolarTermDate   // 該節氣的日期，格式 MM/DD
);
