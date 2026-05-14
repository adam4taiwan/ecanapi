using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models
{
    [Table("BaziTechniques")]
    public class BaziTechnique
    {
        [Key]
        public int Id { get; set; }

        // Category: e.g. 六親-父母, 婚姻, 子女, 事業, 身體, 應期, 神煞, 綜合
        [MaxLength(50)]
        public string Category { get; set; } = "";

        // Search keywords summary
        [MaxLength(200)]
        public string Keywords { get; set; } = "";

        // If-condition (technique trigger)
        public string Condition { get; set; } = "";

        // Then-result (prediction content)
        public string Result { get; set; } = "";
    }
}
