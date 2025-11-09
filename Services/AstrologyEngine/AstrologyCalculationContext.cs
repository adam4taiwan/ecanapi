using Ecan;
using Ecanapi.Models;
using System;
using System.Collections.Generic;

namespace Ecanapi.Services.AstrologyEngine
{
    public class AstrologyCalculationContext
    {
        public AstrologyRequest Request { get; }
        public EcanChineseCalendar Calendar { get; }
        public AstrologyChartResult Result { get; set; }

        public int CUE1 { get; set; }
        public int CUF1 { get; set; }
        public int CUE2 { get; set; }
        public int CUF2 { get; set; }
        public int CUE3 { get; set; }
        public int CUF3 { get; set; }
        public int CUE4 { get; set; }
        public int CUF4 { get; set; }

        public int LunarDay { get; set; }
        public int Day { get; set; }
        public string[] CCN { get; set; } = new string[13];
        public string[] CCM { get; set; } = new string[13];
        public string[] CCB { get; set; } = new string[13];
        public string[] CCO { get; set; } = new string[13];
        public string[] CCX { get; set; } = new string[13];
        public int MingGongIndex { get; set; }
        public int ShenGongIndex { get; set; }
        public int WuXingJu { get; set; }
        public string WuXingJuText { get; set; }
        public string MingZhu { get; set; }
        public string ShenZhu { get; set; }
        public string[] FourTransformationStars { get; set; } = new string[13];
        public string[] SecondaryStars { get; set; } = new string[13];
        public string[] MainStarBrightness { get; set; } = new string[13];
        public string[] PalaceShortNames { get; set; } = new string[13];
        public string[] LifeCycleStage { get; set; } = new string[13];

        public int LuCunPos { get; set; }
        public int ZuoFuPos { get; set; }
        public int YouBiPos { get; set; }
        public int WenChangPos { get; set; }
        public int WenQuPos { get; set; }

        public string[] GoodStars { get; set; } = new string[13];
        public string[] BadStars { get; set; } = new string[13];
        public string[] SmallStars { get; set; } = new string[13];

        public AstrologyCalculationContext(AstrologyRequest request)
        {
            Request = request;
            var birthDate = new DateTime(request.Year, request.Month, request.Day, request.Hour, request.Minute, 0);
            Calendar = new EcanChineseCalendar(birthDate);
            LunarDay = Calendar.ChineseDay;
            Day = request.Day;

            for (int i = 0; i < 13; i++)
            {
                CCN[i] = ""; CCM[i] = ""; CCB[i] = ""; CCO[i] = ""; CCX[i] = "";
                FourTransformationStars[i] = ""; SecondaryStars[i] = "";
                MainStarBrightness[i] = ""; PalaceShortNames[i] = "";
                LifeCycleStage[i] = "";
                GoodStars[i] = ""; BadStars[i] = ""; SmallStars[i] = "";
            }

            var emptyPillar = new PillarInfo("", "", "", "", new List<string>());
            var emptyBazi = new BaziInfo(emptyPillar, emptyPillar, emptyPillar, emptyPillar, "", "");

            Result = new AstrologyChartResult(
                emptyBazi,
                new List<ZiWeiPalace>(),
                "", // WuXingJuText
                "", // MingZhu
                "", // ShenZhu
                new List< BaziLuckCycle > (), // <--- 新增八字大運列表
                null,
                null,
                request.Name,
                birthDate,
                Calendar.ChineseDateString
            );
        }
    }
}