using Ecanapi.Data;
using Ecanapi.Models.Analysis;
using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore;
namespace Ecanapi.Data
{
    // 這裡必須是 class，而不是 interface
    public class ApplicationDbContext : IdentityDbContext<ApplicationUser>
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
            : base(options)
        {
        }

        // 這是您的自定義資料集
        public DbSet<Customer> Customers { get; set; }
        // 【新增】將所有13張分析用表加入到 DbContext
        public DbSet<StarStyle> StarStyles { get; set; }
        public DbSet<PalaceMainStar> PalaceMainStars { get; set; }
        public DbSet<PalaceName> PalaceNames { get; set; }
        public DbSet<PalaceStarBrightness> PalaceStarBrightnesses { get; set; }
        public DbSet<EarthlyBranchHiddenStem> EarthlyBranchHiddenStems { get; set; }
        public DbSet<HeavenlyStemInfo> HeavenlyStemInfos { get; set; }
        public DbSet<NaYin> NaYins { get; set; }
        public DbSet<StarCondition> StarConditions { get; set; }
        public DbSet<BodyMaster> BodyMasters { get; set; }
        public DbSet<WealthOfficialGeneral> WealthOfficialGenerals { get; set; }
        public DbSet<DayPillarToMonthBranch> DayPillarToMonthBranches { get; set; }
        public DbSet<IChing64Hexagrams> IChing64Hexagrams { get; set; }
        public DbSet<IChingExplanation> IChingExplanations { get; set; }
        public DbSet<PreNatalFourTransformations> PreNatalFourTransformations { get; set; }
        public DbSet<PalaceTransformations> PalaceTransformations { get; set; }
        public DbSet<EarthlyBranchStars> EarthlyBranchStars { get; set; }
        public DbSet<HeavenlyStemStars> HeavenlyStemStars { get; set; }
        public DbSet<DayHourStars> DayHourStars { get; set; }
        public DbSet<DayStemToBranch> DayStemToBranches { get; set; }
        public DbSet<SixtyJiaziDayToHour> SixtyJiaziDayToHours { get; set; }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            modelBuilder.Entity<CalendarEntry>()
                .HasKey(c => new { c.Year, c.SolarMonth, c.SolarDay });
        }
    }
}