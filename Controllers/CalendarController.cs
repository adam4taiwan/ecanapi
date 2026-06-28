using Microsoft.AspNetCore.Mvc;
using Ecanapi.Services; // 引入您服務層的命名空間
using Ecanapi.Models;   // 【關鍵修復】加入此行，讓控制器認識 CalendarResponse

namespace Ecanapi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class CalendarController : ControllerBase
    {
        private readonly ICalendarService _calendarService;

        // 透過建構式注入 ICalendarService
        public CalendarController(ICalendarService calendarService)
        {
            _calendarService = calendarService;
        }

        /// <summary>
        /// 根據陽曆日期查詢萬年曆資料
        /// </summary>
        /// <param name="year">西元年</param>
        /// <param name="month">陽曆月份</param>
        /// <param name="day">陽曆日期</param>
        /// <returns>對應日期的日曆資料</returns>
        [HttpGet("{year:int}/{month:int}/{day:int}")]
        [ProducesResponseType(typeof(CalendarResponse), StatusCodes.Status200OK)]
        [ProducesResponseType(StatusCodes.Status404NotFound)]
        public async Task<IActionResult> GetCalendarData(int year, int month, int day)
        {
            var calendarData = await _calendarService.GetCalendarDataAsync(year, month, day);

            if (calendarData == null)
            {
                return NotFound(new { message = $"找不到日期 {year}-{month}-{day} 的資料" });
            }

            // 查詢當天所屬節氣（往前找最近節氣），計算出生日距節氣後第幾天
            string? solarTermInfo = null;
            var prevTerm = await _calendarService.GetPrevSolarTermAsync(year, month, day);
            if (prevTerm != null && !string.IsNullOrEmpty(prevTerm.SolarTerm))
            {
                var termDate = new DateTime(prevTerm.Year, prevTerm.SolarMonth, prevTerm.SolarDay);
                var birthDate = new DateTime(year, month, day);
                int daysDiff = (int)(birthDate - termDate).TotalDays;
                solarTermInfo = daysDiff == 0
                    ? $"{prevTerm.SolarTerm}（當日）"
                    : $"{prevTerm.SolarTerm} 後第{daysDiff}天";
            }

            // 將從資料庫讀取的實體模型，轉換為乾淨的 API 回應模型 (DTO)
            var response = new CalendarResponse(
                calendarData.Year,
                calendarData.SolarMonth,
                calendarData.SolarDay,
                calendarData.YearGanzhi,
                calendarData.MonthGanzhi,
                calendarData.DayGanzhi,
                calendarData.DayTianGan,
                calendarData.SolarTerm,
                calendarData.LunarMonth,
                calendarData.LunarDay,
                calendarData.WeekDay,
                calendarData.Season,
                solarTermInfo
            );

            return Ok(response);
        }
    }
}

