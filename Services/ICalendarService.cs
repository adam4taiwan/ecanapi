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
    }
    #endregion
}
