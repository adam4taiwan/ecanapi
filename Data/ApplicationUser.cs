using Microsoft.AspNetCore.Identity;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Data
{
    // 繼承 IdentityUser 來擴展使用者資料
    public class ApplicationUser : IdentityUser
    {
        public required string Name { get; set; }

        [Column("points")]
        public int Points { get; set; } = 0;

        public ICollection<Customer> Customers { get; set; }

        // ── 命理核心生辰資料（nullable，現有帳號不受影響）──
        public int? BirthYear { get; set; }
        public int? BirthMonth { get; set; }
        public int? BirthDay { get; set; }
        public int? BirthHour { get; set; }
        public int? BirthMinute { get; set; }
        /// <summary>1=男(乾造) 0=女(坤造)</summary>
        public int? BirthGender { get; set; }
        /// <summary>solar=陽曆 lunar=陰曆</summary>
        public string? DateType { get; set; }
        /// <summary>命盤顯示名稱（可與帳號名不同）</summary>
        public string? ChartName { get; set; }

        public bool DeductPoints(int amount)
        {
            if (this.Points < amount) return false;
            this.Points -= amount;
            return true;
        }

        /// <summary>是否已填寫生辰資料</summary>
        public bool HasBirthData =>
            BirthYear.HasValue && BirthMonth.HasValue && BirthDay.HasValue && BirthHour.HasValue;
    }
}
