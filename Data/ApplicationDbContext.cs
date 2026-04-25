using Ecanapi.Controllers;
using Ecanapi.Data;
using Ecanapi.Models;
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
        public DbSet<PointRecord> PointRecords { get; set; }
        public DbSet<DailyFortune> DailyFortunes { get; set; }
        public DbSet<AtmPaymentRequest> AtmPaymentRequests { get; set; }
        public DbSet<FortuneRule> FortuneRules { get; set; }
        public DbSet<KnowledgeDocument> KnowledgeDocuments { get; set; }
        public DbSet<UserChart> UserCharts { get; set; }
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

        // Membership & product catalog
        public DbSet<Product> Products { get; set; }
        public DbSet<MembershipPlan> MembershipPlans { get; set; }
        public DbSet<MembershipPlanBenefit> MembershipPlanBenefits { get; set; }
        public DbSet<UserSubscription> UserSubscriptions { get; set; }
        public DbSet<UserSubscriptionClaim> UserSubscriptionClaims { get; set; }

        // Booking requests (blessing services + consultation appointments)
        public DbSet<BookingRequest> BookingRequests { get; set; }

        // 六十甲子日柱斷語 KB
        public DbSet<BaziDayPillarReading> BaziDayPillarReadings { get; set; }

        // 九星氣學系統
        public DbSet<NineStarTrait> NineStarTraits { get; set; }
        public DbSet<NineStarDailyRule> NineStarDailyRules { get; set; }
        public DbSet<NineStarCombinationRule> NineStarCombinationRules { get; set; }
        public DbSet<LineUser> LineUsers { get; set; }

        // 中原盲派命理秘典 KB
        public DbSet<BaziDirectRule> BaziDirectRules { get; set; }
        public DbSet<FortuneSourceText> FortuneSourceTexts { get; set; }

        // 命書記錄
        public DbSet<UserReport> UserReports { get; set; }

        // 學生白名單
        public DbSet<StudentWhiteList> StudentWhiteLists { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            modelBuilder.Entity<CalendarEntry>()
                .HasKey(c => new { c.Year, c.SolarMonth, c.SolarDay });
        }
    }
}