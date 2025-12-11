using System.Collections.Generic;
using System.Text;

namespace Ecanapi.Models
{
    public class BlindSchoolAnalysisResult
    {
        public string BaZi { get; set; }
        public char RiGan { get; set; }
        public string RiWX { get; set; }
        public TrueFakeInfo TrueFake { get; set; } = new();
        public BodyUseInfo BodyUse { get; set; } = new();
        public PatternInfo Pattern { get; set; } = new();
        public string Conclusion { get; set; }
    }

    public class TrueFakeInfo { public List<ShiShenResult> ShiShens { get; set; } = new(); }
    //public class ShiShenResult
    //{
    //    public char Gan { get; set; }
    //    public string ShiShen { get; set; }
    //    public string WuXing { get; set; }
    //    public bool IsTrue { get; set; }
    //    public bool IsReal { get; set; }
    //}

    public class Ti
    {
        public char Gan { get; set; }
        public string ShiShen { get; set; }
        public string WX { get; set; }
        public string Description { get; set; }
    }

    public class BodyUseInfo
    {
        public string CaiWX { get; set; }
        public string GuanWX { get; set; }
        public List<Ti> Ti { get; set; } = new();
        public string CaiName { get; set; }
    }

    public class PatternInfo
    {
        public string Name { get; set; }
        public string Description { get; set; }
    }
    // 結果類
    public class ShiShenResult
    {
        public char Gan { get; set; }
        public string ShiShen { get; set; }
        public string WuXing { get; set; }
        public bool IsTrueGod { get; set; }
        public bool IsReal { get; set; }
        public string RootType { get; set; }
        public double SitBranchPower { get; set; }
        public List<string> Explanation { get; set; }
        public string Phenomenon { get; set; }

        public override string ToString()
        {
            return $"{ShiShen}({WuXing}) → 真神: {IsTrueGod} | 實: {IsReal}\n" +
                   $"{string.Join("\n", Explanation)}\n" +
                   $"現象：{Phenomenon}\n" +
                   (IsTrueGod && IsReal ? "★★★★★ 真實有力" :
                    !IsTrueGod && !IsReal ? "✘ 100%虛假" :
                    !IsTrueGod ? "★★ 假神虛榮" : "★★★ 真神但虛浮");
        }
    }
}