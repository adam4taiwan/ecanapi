using System;

namespace Ecanapi.Models
{
    public class UserInput
    {
        public required string Name { get; set; }
        public int Gender { get; set; }
        public DateTime BirthDateTime { get; set; }
    }

    public class CalculationResult
    {
        public required string Name { get; set; }
        public int Gender { get; set; }
        public DateTime BirthDateTime { get; set; }
        public string SolarDate { get; set; }
        public string LunarDate { get; set; }
        public string Zodiac { get; set; }
        public string SolarTerm { get; set; }
        public string PreviousSolarTerm { get; set; }
        public string NextSolarTerm { get; set; }
        public string GanZhi { get; set; }
        public string WeekDay { get; set; }
        public string Constellation { get; set; }
        public string ChineseStar { get; set; }
    }
}