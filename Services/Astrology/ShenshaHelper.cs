using System.Collections.Generic;
using System;
using System.Collections.Generic;
using System.IO;
namespace Ecanapi.Services.Astrology
{


    // 基本神煞 (如需更豐富可用 Dictionary+規則法)
    public record LuckShenshaInfo(
        string Name,
        string Desc,
        string RelatedPillar,
        string RelatedValue
    );

    public static class ShenshaHelper
    {
        private static readonly Dictionary<string, LuckShenshaInfo> BaseShenshaTable = new()
        {
            ["太歲"] = new LuckShenshaInfo("太歲", "本年與命盤同支", "全部", "self"),
            // 可自行擴增更多神煞 ("天乙", "驛馬", ...下略)
        };

        public static List<LuckShenshaInfo> CalcYearShensha(
            Dictionary<string, string> natalPillarBranches, // eg: "年"->"寅"
            string yearBranch)
        {
            var result = new List<LuckShenshaInfo>();
            // 太歲，只要命盤任一支與流年同
            if (natalPillarBranches.Values.Contains(yearBranch))
                result.Add(BaseShenshaTable["太歲"]);

            // 更多神煞依日干、出生支、月支等自訂計算
            // if ( ... ) result.Add(BaseShenshaTable["xxx"]);

            return result;
        }

    }

    public class ShenshaRule
    {
        public string Name { get; set; }
        public string Condition { get; set; }
        public string Description { get; set; }
        public string PillarScope { get; set; }
    }

    public static class ShenshaLoaderSimple
    {
        public static List<ShenshaRule> LoadFromCsv(string path, char delimiter = ',')
        {
            var result = new List<ShenshaRule>();
            var lines = File.ReadAllLines(path);
            for (int i = 1; i < lines.Length; i++) // i=1 避開標題row
            {
                var parts = lines[i].Split(delimiter); // 預設 , 可改 \t, ;
                if (parts.Length < 4) continue;
                result.Add(new ShenshaRule
                {
                    Name = parts[0].Trim(),
                    Condition = parts[1].Trim(),
                    Description = parts[2].Trim(),
                    PillarScope = parts[3].Trim()
                });
            }
            return result;
        }
    }

}
