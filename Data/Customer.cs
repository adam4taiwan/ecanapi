namespace Ecanapi.Data
{
    public class Customer
    {
        public int Id { get; set; }
        public required string Name { get; set; }
        public required string Email { get; set; }
        public int Gender { get; set; }
        public DateTime BirthDateTime { get; set; }

        // 唯一客戶編號：由建檔時間 yyyyMMddHHmmss 自動生成
        public string CustomerCode { get; set; } = "";

        // 備注（可選）
        public string? Notes { get; set; }

        // 建檔時間
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

        // 外來鍵，連結到擁有此客戶資料的會員
        public required string ApplicationUserId { get; set; }
        // 修正：將導覽屬性設為可為空值，避免在新增時報錯。
        public ApplicationUser? ApplicationUser { get; set; }
    }
}