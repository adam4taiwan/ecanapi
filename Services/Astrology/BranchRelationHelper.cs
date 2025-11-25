using System;
using System.Collections.Generic;
using System.Linq;
namespace Ecanapi.Services.Astrology
{

    // 關係記錄
    public record BranchRelationInfo(
        string RelationType,       // "合" "沖" "刑" "害" "破"
        string FlowPillar,         // "流年" "流月" "大運"
        string TargetPillar,       // "年" "月" "日" "時" ...
        string FlowBranch,         // 影響源地支
        string TargetBranch,       // 本命目標地支
        string Description         // 中文簡述
    );
    public static class ZhiWuXingHeHelper
    {
        // 六合對照
        public static readonly Dictionary<(string, string), string> ZhiLiuHeTable = new()
    {
        { ("子", "丑"), "土" }, { ("寅", "亥"), "木" }, { ("卯", "戌"), "火" }, { ("辰", "酉"), "金" }, { ("巳", "申"), "水" }, { ("午", "未"), "木" }
    };
        // 各三合局
        public static readonly Dictionary<string[], string> ZhiSanHeTable = new()
    {
        { new[]{ "申", "子", "辰" }, "水" },
        { new[]{ "亥", "卯", "未" }, "木" },
        { new[]{ "寅", "午", "戌" }, "火" },
        { new[]{ "巳", "酉", "丑" }, "金" }
    };
        // 判斷地支合化
        public static string? GetZhiLiuHeWuXing(string zhiA, string zhiB)
        {
            if (ZhiLiuHeTable.TryGetValue((zhiA, zhiB), out var wuxing) ||
                ZhiLiuHeTable.TryGetValue((zhiB, zhiA), out wuxing))
                return wuxing;
            return null;
        }
        // 判斷地支三合（傳入三個地支時用）
        public static string? GetSanHeWuXing(IEnumerable<string> zhiList)
        {
            var arr = zhiList.OrderBy(z => z).ToArray();
            foreach (var group in ZhiSanHeTable.Keys)
            {
                if (group.All(arr.Contains))
                    return ZhiSanHeTable[group];
            }
            return null;
        }
    }

    public static class GanWuXingHeHelper
    {
        // 天干五合結果與合化五行
        public static readonly Dictionary<(string, string), string> GanHeTable = new()
    {
        { ("甲", "己"), "土" }, { ("乙", "庚"), "金" }, { ("丙", "辛"), "水" },
        { ("丁", "壬"), "木" }, { ("戊", "癸"), "火" }
    };

        // 正反都能合
        public static string? GetHeHuaWuXing(string ganA, string ganB)
        {
            if (GanHeTable.TryGetValue((ganA, ganB), out var wuxing) ||
                GanHeTable.TryGetValue((ganB, ganA), out wuxing))
                return wuxing; // 傳回五行
            return null;
        }
    }

    public static class BranchRelationHelper
    {
        // (簡化常用，補充可依需求)
        private static readonly Dictionary<string, string[]> HeTable = new()
    {
        { "子", new[] { "丑" } }, { "丑", new[] { "子" } },
        { "寅", new[] { "亥" } }, { "亥", new[] { "寅" } },
        { "卯", new[] { "戌" } }, { "戌", new[] { "卯" } },
        { "辰", new[] { "酉" } }, { "酉", new[] { "辰" } },
        { "巳", new[] { "申" } }, { "申", new[] { "巳" } },
        { "午", new[] { "未" } }, { "未", new[] { "午" } }
    };
        private static readonly Dictionary<string, string[]> ChongTable = new()
    {
        { "子", new[] { "午" } }, { "午", new[] { "子" } },
        { "丑", new[] { "未" } }, { "未", new[] { "丑" } },
        { "寅", new[] { "申" } }, { "申", new[] { "寅" } },
        { "卯", new[] { "酉" } }, { "酉", new[] { "卯" } },
        { "辰", new[] { "戌" } }, { "戌", new[] { "辰" } },
        { "巳", new[] { "亥" } }, { "亥", new[] { "巳" } }
    };
        private static readonly Dictionary<string, string[]> XingTable = new()
    {
        { "子", new[] { "卯" } }, { "卯", new[] { "子" } },
        { "丑", new[] { "戌" } }, { "戌", new[] { "丑" } },
        { "寅", new[] { "巳" } }, { "巳", new[] { "寅" } },
        { "申", new[] { "寅" } }, { "亥", new[] { "申" } },
        { "未", new[] { "丑" } }
        // 更多三刑補充自加
    };
        private static readonly Dictionary<string, string[]> HaiTable = new()
    {
        { "子", new[] { "未" } }, { "未", new[] { "子" } },
        { "丑", new[] { "午" } }, { "午", new[] { "丑" } },
        { "寅", new[] { "巳" } }, { "巳", new[] { "寅" } },
        { "卯", new[] { "辰" } }, { "辰", new[] { "卯" } },
        { "申", new[] { "亥" } }, { "亥", new[] { "申" } },
        { "酉", new[] { "戌" } }, { "戌", new[] { "酉" } }
    };
        private static readonly Dictionary<string, string[]> PoTable = new()
    {
        { "子", new[] { "酉" } }, { "酉", new[] { "子" } },
        { "午", new[] { "辰" } }, { "辰", new[] { "午" } },
        { "申", new[] { "寅" } }, { "寅", new[] { "申" } },
        { "亥", new[] { "巳" } }, { "巳", new[] { "亥" } }
        // 可補全細項
    };


        public static List<BranchRelationInfo> CalcBranchRelations(
            Dictionary<string, string> natalPillarBranches, // "年"->"寅" "月"->"申"
            string flowBranch,  // 流年支
            string flowPillar = "流年"
        )
        {
            var result = new List<BranchRelationInfo>();
            foreach (var (pillar, natalBranch) in natalPillarBranches)
            {
                if (HeTable.TryGetValue(flowBranch, out var he) && he.Contains(natalBranch))
                    result.Add(new BranchRelationInfo("合", flowPillar, pillar, flowBranch, natalBranch, $"{flowPillar}{flowBranch}與{pillar}{natalBranch}六合"));
                if (ChongTable.TryGetValue(flowBranch, out var chong) && chong.Contains(natalBranch))
                    result.Add(new BranchRelationInfo("沖", flowPillar, pillar, flowBranch, natalBranch, $"{flowPillar}{flowBranch}沖{pillar}{natalBranch}"));
                if (XingTable.TryGetValue(flowBranch, out var xing) && xing.Contains(natalBranch))
                    result.Add(new BranchRelationInfo("刑", flowPillar, pillar, flowBranch, natalBranch, $"{flowPillar}{flowBranch}刑{pillar}{natalBranch}"));
                if (HaiTable.TryGetValue(flowBranch, out var hai) && hai.Contains(natalBranch))
                    result.Add(new BranchRelationInfo("害", flowPillar, pillar, flowBranch, natalBranch, $"{flowPillar}{flowBranch}害{pillar}{natalBranch}"));
                if (PoTable.TryGetValue(flowBranch, out var po) && po.Contains(natalBranch))
                    result.Add(new BranchRelationInfo("破", flowPillar, pillar, flowBranch, natalBranch, $"{flowPillar}{flowBranch}破{pillar}{natalBranch}"));
            }
            return result;
        }
    }


}

