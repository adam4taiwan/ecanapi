using System;

namespace Print2Engine
{
    public interface IEcanCalendar
    {
        string GetChineseDate(DateTime dt);
        string GetChineseZodiac(DateTime dt);
        string GetSolarTerm(DateTime dt);
        string GetPreviousSolarTerm(DateTime dt);
        string GetNextSolarTerm(DateTime dt);
        string GetGanZhi(DateTime dt);
        string GetConstellation(DateTime dt);
        string GetChineseConstellation(DateTime dt);
    }
}