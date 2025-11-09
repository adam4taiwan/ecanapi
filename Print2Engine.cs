using System;
using Ecanapi.Models;
namespace Print2Engine;

public sealed class Print2Engine
{
    private readonly IEcanCalendar _calendar;

    public Print2Engine(IEcanCalendar calendar)
    {
        _calendar = calendar ?? throw new ArgumentNullException(nameof(calendar));
    }

    public CalculationResult Process(UserInput input)
    {
        if (input == null) throw new ArgumentNullException(nameof(input));
        var dt = input.BirthDateTime;

        var result = new CalculationResult
        {
            Name = input.Name,
            Gender = input.Gender,
            BirthDateTime = dt,

            SolarDate       = dt.ToString("yyyy-MM-dd HH:mm"),
            LunarDate       = _calendar.GetChineseDate(dt),
            Zodiac          = _calendar.GetChineseZodiac(dt),
            SolarTerm       = _calendar.GetSolarTerm(dt),
            PreviousSolarTerm = _calendar.GetPreviousSolarTerm(dt),
            NextSolarTerm   = _calendar.GetNextSolarTerm(dt),
            GanZhi          = _calendar.GetGanZhi(dt),
            WeekDay         = dt.DayOfWeek.ToString(),
            Constellation   = _calendar.GetConstellation(dt),
            ChineseStar     = _calendar.GetChineseConstellation(dt)
        };

        return result;
    }
}
