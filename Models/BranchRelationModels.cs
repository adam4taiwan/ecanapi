namespace Ecanapi.Models
{
    public class BranchRelationModels
    {
        public record BranchRelationInfo(
            string RelationType,       // 如 合、沖、刑、害、破
            string FlowPillar,         // 影響源：流年、流月、大運
            string TargetPillar,       // 被影響柱：年/月/日/時
            string FlowBranch,         // 流年地支
            string TargetBranch,       // 本命四柱地支
            string Description         // 中文描述或重點
        );

        public record LuckShenshaInfo(
            string Name,         // 神煞名稱
            string Desc,         // 中文短描述
            string RelatedPillar,// 有效的柱位類別
            string RelatedValue  // 地支或天干
        );
        public static class BaZiBranchRelation
        {
            // 八字地支合、沖、刑、害、破 對照表
            // 合
            private static readonly Dictionary<string, string[]> HeTable = new()
            {
                ["子"] = new[] { "丑" },
                ["丑"] = new[] { "子" },
                ["寅"] = new[] { "亥" },
                ["亥"] = new[] { "寅" },
                ["卯"] = new[] { "戌" },
                ["戌"] = new[] { "卯" },
                ["辰"] = new[] { "酉" },
                ["酉"] = new[] { "辰" },
                ["巳"] = new[] { "申" },
                ["申"] = new[] { "巳" },
                ["午"] = new[] { "未" },
                ["未"] = new[] { "午" }
            };
            // 沖
            private static readonly Dictionary<string, string[]> ChongTable = new()
            {
                ["子"] = new[] { "午" },
                ["午"] = new[] { "子" },
                ["丑"] = new[] { "未" },
                ["未"] = new[] { "丑" },
                ["寅"] = new[] { "申" },
                ["申"] = new[] { "寅" },
                ["卯"] = new[] { "酉" },
                ["酉"] = new[] { "卯" },
                ["辰"] = new[] { "戌" },
                ["戌"] = new[] { "辰" },
                ["巳"] = new[] { "亥" },
                ["亥"] = new[] { "巳" }
            };
            // 刑
            private static readonly Dictionary<string, string[]> XingTable = new()
            {
                ["子"] = new[] { "卯" },
                ["卯"] = new[] { "子" },
                ["丑"] = new[] { "戌" },
                ["戌"] = new[] { "丑" },
                ["寅"] = new[] { "巳" },
                ["巳"] = new[] { "寅" },
                ["申"] = new[] { "寅" },
                ["亥"] = new[] { "申" },
                ["未"] = new[] { "丑" }
                // 省略複雜的自刑、三刑可依真正需求補充
            };
            // 害
            private static readonly Dictionary<string, string[]> HaiTable = new()
            {
                ["子"] = new[] { "未" },
                ["未"] = new[] { "子" },
                ["丑"] = new[] { "午" },
                ["午"] = new[] { "丑" },
                ["寅"] = new[] { "巳" },
                ["巳"] = new[] { "寅" },
                ["卯"] = new[] { "辰" },
                ["辰"] = new[] { "卯" },
                ["申"] = new[] { "亥" },
                ["亥"] = new[] { "申" },
                ["酉"] = new[] { "戌" },
                ["戌"] = new[] { "酉" }
            };
            // 破
            private static readonly Dictionary<string, string[]> PoTable = new()
            {
                ["子"] = new[] { "酉" },
                ["酉"] = new[] { "子" },
                ["午"] = new[] { "辰" },
                ["辰"] = new[] { "午" },
                ["申"] = new[] { "寅" },
                ["寅"] = new[] { "申" },
                ["亥"] = new[] { "巳" },
                ["巳"] = new[] { "亥" },
                // 其他特殊破可擴充
            };

            public static List<BranchRelationInfo> CalcBranchRelations(
                Dictionary<string, string> natalPillarBranches, // "年"->"巳"
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
}
