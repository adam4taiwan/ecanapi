using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models
{
    // Master product catalog - every purchasable item has a code
    public class Product
    {
        public int Id { get; set; }

        [Required, MaxLength(50)]
        public string Code { get; set; } = "";        // e.g. BOOK_BAZI, BLESSING_ANTAISUI, COURSE_BASIC

        [Required, MaxLength(100)]
        public string Name { get; set; } = "";        // display name

        [Required, MaxLength(30)]
        public string Type { get; set; } = "";        // book / blessing / consultation / course / lecture / daily / subscription

        public int? PointCost { get; set; }           // cost in points (null = not purchasable with points)
        public int? PriceTwd { get; set; }            // direct price in TWD (null = not directly purchasable)
        public bool IsActive { get; set; } = true;
        public int SortOrder { get; set; }
        public string? Description { get; set; }
        public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
    }

    // Subscription plan definitions (Gold / Silver / Bronze)
    public class MembershipPlan
    {
        public int Id { get; set; }

        [Required, MaxLength(20)]
        public string Code { get; set; } = "";        // GOLD / SILVER / BRONZE

        [Required, MaxLength(50)]
        public string Name { get; set; } = "";        // display name

        public int PriceTwd { get; set; }
        public int DurationDays { get; set; } = 365;
        public bool IsActive { get; set; } = true;
        public int SortOrder { get; set; }
        public string? Description { get; set; }

        public ICollection<MembershipPlanBenefit> Benefits { get; set; } = new List<MembershipPlanBenefit>();
    }

    // Benefits included in each plan - flexible composition
    public class MembershipPlanBenefit
    {
        public int Id { get; set; }
        public int PlanId { get; set; }
        public MembershipPlan Plan { get; set; } = null!;

        // Target: either a specific product code or an entire product type
        [MaxLength(50)]
        public string? ProductCode { get; set; }      // e.g. BOOK_LIUNIAN (specific product)

        [MaxLength(30)]
        public string? ProductType { get; set; }      // e.g. book (entire category)

        // What benefit is granted
        [Required, MaxLength(20)]
        public string BenefitType { get; set; } = ""; // free / discount / quota / access

        [Required, MaxLength(20)]
        public string BenefitValue { get; set; } = ""; // "1" = 1 free item; "0.85" = 85% price; "true" = access

        public string? Description { get; set; }
    }

    // Tracks each user's active or past subscriptions
    public class UserSubscription
    {
        public int Id { get; set; }

        [Required]
        public string UserId { get; set; } = "";

        public int PlanId { get; set; }
        public MembershipPlan Plan { get; set; } = null!;

        public DateTime StartDate { get; set; }
        public DateTime ExpiryDate { get; set; }

        [MaxLength(20)]
        public string Status { get; set; } = "active"; // active / expired / cancelled

        [MaxLength(100)]
        public string? PaymentRef { get; set; }        // ECPay MerchantTradeNo

        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }

    // Tracks consumption of quota-based benefits (e.g. 1 free liunian book per year)
    public class UserSubscriptionClaim
    {
        public int Id { get; set; }

        [Required]
        public string UserId { get; set; } = "";

        public int SubscriptionId { get; set; }

        [Required, MaxLength(50)]
        public string ProductCode { get; set; } = "";

        public int? ClaimYear { get; set; }            // for annual quota checks
        public DateTime ClaimedAt { get; set; } = DateTime.UtcNow;
    }

    // Records booking requests for blessing services and consultation appointments
    public class BookingRequest
    {
        public int Id { get; set; }

        public string? UserId { get; set; }            // null if submitted without login

        [Required, MaxLength(20)]
        public string ServiceType { get; set; } = ""; // blessing / consultation

        [MaxLength(50)]
        public string? ServiceCode { get; set; }       // BLESSING_ANTAISUI / BLESSING_LIGHT / etc.

        [Required, MaxLength(100)]
        public string Name { get; set; } = "";

        [MaxLength(200)]
        public string? ContactInfo { get; set; }       // LINE ID / WeChat / phone number

        [MaxLength(20)]
        public string? ContactType { get; set; }       // line / wechat / phone

        [MaxLength(20)]
        public string? BirthDate { get; set; }         // free-form date string

        public bool IsLunar { get; set; } = false;     // whether BirthDate is lunar calendar

        [MaxLength(500)]
        public string? Notes { get; set; }             // user's message / question topic

        [MaxLength(50)]
        public string? PreferredDate { get; set; }     // preferred appointment date range

        [MaxLength(20)]
        public string Status { get; set; } = "pending"; // pending / confirmed / completed / cancelled

        [MaxLength(200)]
        public string? AdminNote { get; set; }

        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
    }
}
