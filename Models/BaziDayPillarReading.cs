using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ecanapi.Models
{
    [Table("BaziDayPillarReadings")]
    public class BaziDayPillarReading
    {
        [Key]
        public int Id { get; set; }
        public int Sequence { get; set; }

        [MaxLength(2)]
        public string DayPillar { get; set; } = "";

        [MaxLength(1)]
        public string DayStem { get; set; } = "";

        [MaxLength(1)]
        public string DayBranch { get; set; } = "";

        public string? Overview { get; set; }
        public string? ShenAnalysis { get; set; }
        public string? InnerTraits { get; set; }
        public string? Career { get; set; }
        public string? Weaknesses { get; set; }
        public string? MotherInfluence { get; set; }
        public string? FatherInfluence { get; set; }
        public string? SiblingInfluence { get; set; }
        public string? ChildInfluence { get; set; }
        public string? MonthInfluence { get; set; }
        public string? MaleChart { get; set; }
        public string? FemaleChart { get; set; }
        public string? SpecialHours { get; set; }
        public DateTime CreatedAt { get; set; }
    }
}
