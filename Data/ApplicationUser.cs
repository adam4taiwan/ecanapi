using Microsoft.AspNetCore.Identity;

namespace Ecanapi.Data
{
    // 繼承 IdentityUser 來擴展使用者資料
    public class ApplicationUser : IdentityUser
    {
        // 新增一個屬性來儲存使用者的姓名
        public required string Name { get; set; }

        // 一個導覽屬性，連結到此使用者擁有的所有客戶
        public ICollection<Customer> Customers { get; set; }
    }
}
