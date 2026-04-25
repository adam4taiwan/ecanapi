using System.ComponentModel.DataAnnotations;

namespace Ecanapi.Models
{
    public class StudentWhiteList
    {
        public int Id { get; set; }

        [Required, MaxLength(255)]
        public string Email { get; set; } = string.Empty;

        [MaxLength(255)]
        public string? Note { get; set; }

        [MaxLength(255)]
        public string AddedByEmail { get; set; } = string.Empty;

        public DateTime AddedAt { get; set; } = DateTime.UtcNow;
    }
}
