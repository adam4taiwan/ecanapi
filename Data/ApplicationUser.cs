using Microsoft.AspNetCore.Identity;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Data
{
    // 繼承 IdentityUser 來擴展使用者資料
    public class ApplicationUser : IdentityUser
    {
        // 新增一個屬性來儲存使用者的姓名
        public required string Name { get; set; }

        // Points 是新加的，對應資料庫您剛才 ALTER 的欄位
        [Column("points")]
        public int Points { get; set; } = 0;
        // 一個導覽屬性，連結到此使用者擁有的所有客戶
        public ICollection<Customer> Customers { get; set; }

        public bool DeductPoints(int amount)
        {
            if (this.Points < amount) return false;
            this.Points -= amount;
            return true;
        }
    }
}
