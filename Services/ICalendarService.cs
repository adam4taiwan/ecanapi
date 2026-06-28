using Ecanapi.Data;
using Ecanapi.Models.Analysis;
using Microsoft.EntityFrameworkCore;

namespace Ecanapi.Services
{
    #region 服務層 (Service Layer)
    /// <summary>
    /// 定義日曆服務的介面。
    /// </summary>
    public interface ICalendarService
    {
        Task<CalendarEntry?> GetCalendarDataAsync(int year, int month, int day);
        Task<CalendarEntry?> GetPrevSolarTermAsync(int year, int month, int day);
    }

    /// <summary>
    /// 實作日曆服務，處理資料庫查詢邏輯。
    /// </summary>
    public class CalendarService : ICalendarService
    {
        private readonly CalendarDbContext _context;

        public CalendarService(CalendarDbContext context)
        {
            _context = context;
        }

        public async Task<CalendarEntry?> GetCalendarDataAsync(int year, int month, int day)
        {
            // FirstOrDefaultAsync 方法現在可以被正確識別了，因為它來自 Microsoft.EntityFrameworkCore
            return await _context.CalendarEntries
                .FirstOrDefaultAsync(c => c.Year == year && c.SolarMonth == month && c.SolarDay == day);
        }

        public async Task<CalendarEntry?> GetPrevSolarTermAsync(int year, int month, int day)
        {
            // 查詢在指定日期當天或之前、最近一個有節氣的日曆條目
            // 必須用 FromSqlInterpolated，EF Core LINQ 對 calendar 表有回傳 null 的 bug
            int dateSerial = year * 10000 + month * 100 + day;
            var results = await _context.CalendarEntries
                .FromSqlInterpolated($@"
                    SELECT * FROM public.calendar
                    WHERE (""西元年"" * 10000 + ""陽月"" * 100 + ""陽日"") <= {dateSerial}
                    AND ""節氣"" IS NOT NULL AND ""節氣"" != ''
                    ORDER BY ""西元年"" DESC, ""陽月"" DESC, ""陽日"" DESC
                    LIMIT 1")
                .ToListAsync();
            return results.FirstOrDefault();
        }
    }
    #endregion
}
