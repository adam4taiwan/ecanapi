using Microsoft.EntityFrameworkCore;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Data; // 確認命名空間與您的專案結構一致

/// <summary>
/// 這個 DbContext 類別是 Entity Framework Core 用來與資料庫溝通的核心。
/// </summary>
public class CalendarDbContext : DbContext
{
    public CalendarDbContext(DbContextOptions<CalendarDbContext> options) : base(options) { }

    // 這個 DbSet<CalendarEntry> 就代表了您的 public.calendar 資料表
    public DbSet<CalendarEntry> CalendarEntries { get; set; }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        // 由於原始資料表沒有主鍵，我們在此手動將「西元年、陽月、陽日」設定為複合主鍵。
        // 這對於 EF Core 正常運作是必要的。
        modelBuilder.Entity<CalendarEntry>()
            .HasKey(c => new { c.Year, c.SolarMonth, c.SolarDay });
    }
}


/// <summary>
/// 此類別的屬性與 PostgreSQL 中 `public.calendar` 資料表的欄位一一對應。
/// 這個類別可以放在同一個檔案，也可以獨立成 Models/Calendar.cs
/// 從您的截圖看，您似乎已在 Models/calendar.cs 建立了它，請確認內容是否如下。
/// </summary>
[Table("calendar", Schema = "public")]
public class CalendarEntry
{
    [Column("西元年")]
    public int Year { get; set; }

    [Column("陽月")]
    public int SolarMonth { get; set; }

    [Column("陽日")]
    public int SolarDay { get; set; }

    [Column("年干支")]
    public string? YearGanzhi { get; set; }

    [Column("月干支")]
    public string? MonthGanzhi { get; set; }

    [Column("日干支")]
    public string? DayGanzhi { get; set; }

    [Column("日天干")]
    public string? DayTianGan { get; set; }

    [Column("節氣")]
    public string? SolarTerm { get; set; }

    [Column("陰曆月")]
    public string? LunarMonth { get; set; }

    [Column("陰曆日")]
    public string? LunarDay { get; set; }

    [Column("星期")]
    public string? WeekDay { get; set; }

    [Column("季節")]
    public string? Season { get; set; }
}
