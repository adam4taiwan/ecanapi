namespace Ecanapi.Models
{
    public class AtmPaymentRequest
    {
        public int Id { get; set; }
        public string UserId { get; set; } = string.Empty;
        public string PackageId { get; set; } = string.Empty;
        public int Points { get; set; }
        public int PriceTwd { get; set; }
        public string TransferDate { get; set; } = string.Empty;
        public string AccountLast5 { get; set; } = string.Empty;
        public string Status { get; set; } = "pending";
        public string? AdminNote { get; set; }
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        public DateTime? ProcessedAt { get; set; }
    }
}
