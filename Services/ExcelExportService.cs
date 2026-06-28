using Ecanapi.Models;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ecanapi.Services
{
    public class ExcelExportService : IExcelExportService
    {
        private class PalaceCoords
        {
            public (int Row, int Col) PalaceName { get; set; }
            public (int Row, int Col) DecadeAge { get; set; }
            public List<(int Row, int Col)> MainStars { get; set; } = new List<(int, int)>();
            public List<(int Row, int Col)> Brightness { get; set; } = new List<(int, int)>();
            public (int Row, int Col) CombinedStars { get; set; }

            public List<(int Row, int Col)> AnnualTrans { get; set; } = new List<(int, int)>();
            // public (int Row, int Col) AnnualTrans { get; set; }
            public (int Row, int Col) DoctorStar { get; set; }
            public (int Row, int Col) OverflowStars { get; set; } // stars 6+ overflow here (DoctorStar col+1)
            public (int Row, int Col) CombinedInfo { get; set; }
            public (int Row, int Col) EarthlyBranch { get; set; }
        }


        public async Task<byte[]> GenerateChartAsync(AstrologyChartResult chartData, string templatePath)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            await using var fileStream = new FileStream(templatePath, FileMode.Open, FileAccess.Read);
            var workbook = new HSSFWorkbook(fileStream);
            ISheet sheet = workbook.GetSheetAt(0);

            // 1. 填寫盤首資訊
            SetCellValue(sheet, 5, 8, chartData.UserName);
            string birthDateText = chartData.SolarBirthDate.ToString("yyyy-MM-dd HH:mm");
            if (!string.IsNullOrEmpty(chartData.SolarTermInfo))
                birthDateText += "  節氣：" + chartData.SolarTermInfo;
            SetCellValue(sheet, 46, 7, birthDateText);
            SetCellValue(sheet, 5, 2, chartData.LunarBirthDate);
            SetCellValue(sheet, 5, 17, chartData.WuXingJuText);
            SetCellValue(sheet, 5, 7, "命主" + chartData.MingZhu);
            SetCellValue(sheet, 5, 12, "身主" + chartData.ShenZhu);

     
            // 日干十二長生對照表（日干 → 地支順序從長生開始）
            var changShengMap = new Dictionary<string, string[]>
            {
                { "甲", new[]{"亥","子","丑","寅","卯","辰","巳","午","未","申","酉","戌"} },
                { "乙", new[]{"午","巳","辰","卯","寅","丑","子","亥","戌","酉","申","未"} },
                { "丙", new[]{"寅","卯","辰","巳","午","未","申","酉","戌","亥","子","丑"} },
                { "丁", new[]{"酉","申","未","午","巳","辰","卯","寅","丑","子","亥","戌"} },
                { "戊", new[]{"寅","卯","辰","巳","午","未","申","酉","戌","亥","子","丑"} },
                { "己", new[]{"酉","申","未","午","巳","辰","卯","寅","丑","子","亥","戌"} },
                { "庚", new[]{"巳","午","未","申","酉","戌","亥","子","丑","寅","卯","辰"} },
                { "辛", new[]{"子","亥","戌","酉","申","未","午","巳","辰","卯","寅","丑"} },
                { "壬", new[]{"申","酉","戌","亥","子","丑","寅","卯","辰","巳","午","未"} },
                { "癸", new[]{"卯","寅","丑","子","亥","戌","酉","申","未","午","巳","辰"} },
            };
            var changShengNames = new[] { "長生","沐浴","冠帶","臨官","帝旺","衰","病","死","墓","絕","胎","養" };

            string dayStem = chartData.Bazi?.DayPillar?.HeavenlyStem ?? "";

            // 八字大運
            for (int i = 0; i < chartData.BaziLuckCycles.Count; i++)
            {
                SetCellValueBlack(sheet, 31, 13 - i, chartData.BaziLuckCycles[i].StartAge.ToString());
                SetCellValueBlack(sheet, 32, 13 - i, chartData.BaziLuckCycles[i].LiuShen);
                SetCellValueBlack(sheet, 33, 13 - i, chartData.BaziLuckCycles[i].HeavenlyStem);
                SetCellValueBlack(sheet, 34, 13 - i, chartData.BaziLuckCycles[i].EarthlyBranch);

                // 十二長生
                if (changShengMap.TryGetValue(dayStem, out var branchOrder))
                {
                    string branch = chartData.BaziLuckCycles[i].EarthlyBranch;
                    int idx = Array.IndexOf(branchOrder, branch);
                    if (idx >= 0)
                        SetCellValueBlack(sheet, 35, 13 - i, changShengNames[idx]);
                }
            }

   


 

            //new
            // ---------------------------
            // Helper: 寫入 Pillar（包含藏干 pair）
            // ---------------------------
            void FillPillarSafe(int col, PillarInfo pillar)
            {
                if (pillar == null) return;
                // 直接對應 JSON 欄位（確保和 swagger 一致）
                SetCellValue(sheet, 18, col, pillar.HeavenlyStemLiuShen ?? "");
                SetCellValue(sheet, 19, col, pillar.HeavenlyStem ?? "");
                SetCellValue(sheet, 24, col, pillar.EarthlyBranch ?? "");
                SetCellValue(sheet, 27, col, pillar.NaYin ?? "");

                // HiddenStemLiuShen 是交錯放 (liuShen, stem, liuShen, stem...)
                var hidden = pillar.HiddenStemLiuShen ?? new List<string>();
                // 模板放置位置：從 row 28 開始，每兩個為一列 (28/29/30...), col 與 col+1 配對
                int pairIndex = 0;
                for (int i = 0; i + 1 < hidden.Count; i += 2)
                {
                    int writeRow = 28 + pairIndex;       // 28,29,30...
                    int writeCol1 = col;                 // liuShen 放在 col
                    int writeCol2 = col + 1;             // stem 放在 col+1
                    string liu = hidden[i] ?? "";
                    string stem = hidden[i + 1] ?? "";
                    SetCellValue(sheet, writeRow, writeCol1, liu);
                    SetCellValue(sheet, writeRow, writeCol2, stem);
                    pairIndex++;
                }
                // 如果有單數筆（理論上不會，但防呆）
                if (hidden.Count % 2 == 1)
                {
                    int writeRow = 28 + pairIndex;
                    SetCellValue(sheet, writeRow, col, hidden[^1] ?? "");
                }
            }

            // 呼叫（年/月/日/時)
            FillPillarSafe(12, chartData.Bazi.YearPillar);
            FillPillarSafe(10, chartData.Bazi.MonthPillar);
            FillPillarSafe(8, chartData.Bazi.DayPillar);
            FillPillarSafe(6, chartData.Bazi.TimePillar);

            // 3. 建立十二宮座標地圖 (根據您 `正確.xlsx` 檔案完全重製)
            var coordsMap = new Dictionary<int, PalaceCoords>
            {
                { 1, new PalaceCoords { PalaceName = (36, 12), DecadeAge = (36, 10), MainStars = new List<(int,int)>{(36, 13), (36, 14)}, Brightness = new List<(int,int)>{(38,13), (38,14)}, CombinedStars = (37, 10), AnnualTrans = new List<(int,int)>{(39,13), (39,14)}, DoctorStar = (39,10), OverflowStars = (40,10), CombinedInfo = (44, 10),EarthlyBranch = (45, 10) } }, // 子
                { 2, new PalaceCoords { PalaceName = (36, 7), DecadeAge = (36, 5), MainStars = new List<(int,int)>{(36, 8), (36, 9)}, Brightness = new List<(int,int)>{(38,8), (38,9)}, CombinedStars = (37, 5), AnnualTrans = new List<(int,int)>{(39,8), (39,9)}, DoctorStar = (39,5), OverflowStars = (40,5), CombinedInfo = (44, 5),EarthlyBranch = (45, 5) } }, // 丑
                { 3, new PalaceCoords { PalaceName = (36, 2), DecadeAge = (36, 0), MainStars = new List<(int,int)>{(36, 3), (36, 4)}, Brightness = new List<(int,int)>{(38,3), (38,4)}, CombinedStars = (37, 0), AnnualTrans = new List<(int,int)>{(39,3), (39,4)}, DoctorStar = (39,0), OverflowStars = (40,0), CombinedInfo = (44, 0),EarthlyBranch = (45, 0) } }, // 寅
                { 4, new PalaceCoords { PalaceName = (26, 2), DecadeAge = (26, 0), MainStars = new List<(int,int)>{(26, 3), (26, 4)}, Brightness = new List<(int,int)>{(28,3), (28,4)}, CombinedStars = (27, 0), AnnualTrans = new List<(int,int)>{(29,3), (29,4)}, DoctorStar = (29,0), OverflowStars = (30,0), CombinedInfo = (34, 0),EarthlyBranch = (35, 0) } }, // 卯
                { 5, new PalaceCoords { PalaceName = (16, 2), DecadeAge = (16, 0), MainStars = new List<(int,int)>{(16, 3), (16, 4)}, Brightness = new List<(int,int)>{(18,3), (18,4)}, CombinedStars = (17, 0), AnnualTrans = new List<(int,int)>{(19,3), (19,4)}, DoctorStar = (19,0), OverflowStars = (20,0), CombinedInfo = (24, 0),EarthlyBranch = (25, 0) } }, // 辰
                { 6, new PalaceCoords { PalaceName = (6, 2), DecadeAge = (6, 0), MainStars = new List<(int,int)>{(6, 3), (6, 4)}, Brightness = new List<(int,int)>{(8,3), (8,4)}, CombinedStars = (7, 0), AnnualTrans = new List<(int,int)>{(9,3), (9,4)}, DoctorStar = (9,0), OverflowStars = (10,0), CombinedInfo = (14, 0),EarthlyBranch = (15, 0) } }, // 巳
                { 7, new PalaceCoords { PalaceName = (6, 7), DecadeAge = (6, 5), MainStars = new List<(int,int)>{(6, 8), (6, 9)}, Brightness = new List<(int,int)>{(8,8), (8,9)}, CombinedStars = (7, 5), AnnualTrans = new List<(int,int)>{(9,8), (9,9)}, DoctorStar = (9,5), OverflowStars = (10,5), CombinedInfo = (14, 5),EarthlyBranch = (15, 5) } }, // 午
                { 8, new PalaceCoords { PalaceName = (6, 12), DecadeAge = (6, 10), MainStars = new List<(int,int)>{(6, 13), (6, 14)}, Brightness = new List<(int,int)>{(8,13),(8,14)}, CombinedStars = (7, 10), AnnualTrans = new List<(int,int)>{(9,13), (9,14)}, DoctorStar = (9,10), OverflowStars = (10,10), CombinedInfo = (14, 10),EarthlyBranch = (15, 10) } }, // 未
                { 9, new PalaceCoords { PalaceName = (6, 17), DecadeAge = (6, 15), MainStars = new List<(int,int)>{(6, 18), (6, 19)}, Brightness = new List<(int,int)>{(8,18),(8,19)}, CombinedStars = (7, 15), AnnualTrans = new List<(int,int)>{(9,18), (9,19)}, DoctorStar = (9,15), OverflowStars = (10,15), CombinedInfo = (14, 15),EarthlyBranch = (15, 15) } }, // 申
                { 10, new PalaceCoords { PalaceName = (16, 17), DecadeAge = (16, 15), MainStars = new List<(int,int)>{(16, 18), (16, 19)}, Brightness = new List<(int,int)>{(18,18),(18,19)}, CombinedStars = (17, 15), AnnualTrans = new List<(int,int)>{(19,18), (19,19)}, DoctorStar = (19,15), OverflowStars = (20,15), CombinedInfo = (24, 15),EarthlyBranch = (25, 15) } }, // 酉
                { 11, new PalaceCoords { PalaceName = (26, 17), DecadeAge = (26, 15), MainStars = new List<(int,int)>{(26, 18), (26, 19)}, Brightness = new List<(int,int)>{(28,18),(28,19)}, CombinedStars = (27, 15), AnnualTrans = new List<(int,int)>{(29,18), (29,19)}, DoctorStar = (29,15), OverflowStars = (30,15), CombinedInfo = (34, 15),EarthlyBranch = (35, 15) } }, // 戌
                { 12, new PalaceCoords { PalaceName = (36, 17), DecadeAge = (36, 15), MainStars = new List<(int,int)>{(36, 18), (36, 19)}, Brightness = new List<(int,int)>{(38,18),(38,19)}, CombinedStars = (37, 15), AnnualTrans = new List<(int,int)>{(39,18), (39,19)}, DoctorStar = (39,15), OverflowStars = (40,15), CombinedInfo = (44, 15),EarthlyBranch = (45, 15) } } // 亥
            };

            foreach (var palace in chartData.palaces)
            {
                if (coordsMap.TryGetValue(palace.Index, out var coords))
                {
                    SetCellValue(sheet, coords.PalaceName.Row, coords.PalaceName.Col, palace.PalaceName);
                    // 大運
                    SetCellValue(sheet, coords.DecadeAge.Row, coords.DecadeAge.Col, palace.DecadeAgeRange);
                    // 主星
                    for (int i = 0; i < palace.MajorStars.Count; i++)
                    {
                        if (i < coords.MainStars.Count)
                            SetCellValue(sheet, coords.MainStars[i].Row, coords.MainStars[i].Col, palace.MajorStars[i]);
                    }

                    var brightnessValues = palace.MainStarBrightness.Split(',');
                    for (int i = 0; i < brightnessValues.Length; i++)
                    {
                        if (i < coords.Brightness.Count)
                            SetCellValue(sheet, coords.Brightness[i].Row, coords.Brightness[i].Col, brightnessValues[i]);
                    }

                    //var allAuxStars = palace.SecondaryStars.Concat(palace.GoodStars).Concat(palace.BadStars);
                    //SetCellValue(sheet, coords.CombinedStars.Row, coords.CombinedStars.Col, string.Join(" ", allAuxStars));
                    // === 修正版：清理並按預期順序合併副星 ===
                    var secondary = palace.SecondaryStars ?? new List<string>();
                    var good = palace.GoodStars ?? new List<string>();
                    var bad = palace.BadStars ?? new List<string>();
                    //var smalls  = palace.SmallStars ?? new List<string>();

                    // 收集所有副星（去重、去空），同時提取非主星四化標注（如 "曲(忌)" -> "曲" + "忌"）
                    List<string> allAuxStars = new List<string>();
                    var auxTransChars = new List<string>();
                    foreach (var src in new[] { secondary, good, bad })
                        foreach (var s in src)
                        {
                            if (string.IsNullOrWhiteSpace(s)) continue;
                            var m = System.Text.RegularExpressions.Regex.Match(s, @"^(.+)\((.+)\)$");
                            if (m.Success)
                            {
                                var cleanName = m.Groups[1].Value;
                                if (!allAuxStars.Contains(cleanName)) allAuxStars.Add(cleanName);
                                auxTransChars.Add(m.Groups[2].Value);
                            }
                            else
                            {
                                if (!allAuxStars.Contains(s)) allAuxStars.Add(s);
                            }
                        }

                    // 依重要性排序：P1六吉星 > P2四煞 > P3空亡 > P4雜星
                    var p1 = new HashSet<string> { "左", "右", "昌", "曲", "魁", "鉞" };
                    var p2 = new HashSet<string> { "陀", "羊", "火", "鈴" };
                    var p3 = new HashSet<string> { "旬", "截", "空", "劫" };
                    int StarPriority(string s) {
                        var c = s.Length > 0 ? s[..1] : "";
                        if (p1.Contains(c)) return 1;
                        if (p2.Contains(c)) return 2;
                        if (p3.Contains(c)) return 3;
                        return 4;
                    }
                    var sorted = allAuxStars.OrderBy(StarPriority).ToList();

                    // P1+P2+P3 全部進 CombinedStars，P4 雜星補到 5 格，其餘溢出至博士星旁
                    const int MaxP4InMain = 5;
                    var mainDisplay = sorted.Where(s => StarPriority(s) < 4).ToList();
                    var p4Stars = sorted.Where(s => StarPriority(s) == 4).ToList();
                    int p4Slots = Math.Max(0, MaxP4InMain - mainDisplay.Count);
                    mainDisplay.AddRange(p4Stars.Take(p4Slots));
                    var overflowDisplay = p4Stars.Skip(p4Slots).ToList();

                    SetCellValue(sheet, coords.CombinedStars.Row, coords.CombinedStars.Col, string.Join(" ", mainDisplay).Trim());
                    // 非主星四化（如 昌/曲 化忌）顯示在 CombinedStars 下方同欄
                    if (auxTransChars.Count > 0)
                        SetCellValue(sheet, coords.CombinedStars.Row + 1, coords.CombinedStars.Col, string.Join(" ", auxTransChars));

                    // 主星四化：與 MajorStars 同步 compact，index 對齊 AnnualTrans 欄位
                    for (int i = 0; i < palace.AnnualStarTransformations.Count; i++)
                    {
                        if (i < coords.AnnualTrans.Count)
                            SetCellValue(sheet, coords.AnnualTrans[i].Row, coords.AnnualTrans[i].Col, palace.AnnualStarTransformations[i]);
                    }
                    //SetCellValue(sheet, coords.AnnualTrans[i].Row, coords.AnnualTrans[i].Col, string.Join("", palace.AnnualStarTransformations));

                    var smallStarParts = palace.SmallStars;
                    string doctorStar = smallStarParts.Count > 0 ? smallStarParts[0] : "";
                    string ageStar = smallStarParts.Count > 1 ? smallStarParts[1] : "";
                    string generalStar = smallStarParts.Count > 2 ? smallStarParts[2] : "";
                    //
                    SetCellValue(sheet, coords.DoctorStar.Row, coords.DoctorStar.Col, doctorStar);
                    if (overflowDisplay.Count > 0)
                        SetCellValue(sheet, coords.OverflowStars.Row, coords.OverflowStars.Col, string.Join(" ", overflowDisplay));
                    //
                    string combinedInfoString = $"{palace.PalaceStem}-{palace.PalaceStemTransformations} {ageStar}{generalStar} {palace.LifeCycleStage}";
                    SetCellValue(sheet, coords.CombinedInfo.Row, coords.CombinedInfo.Col, combinedInfoString);
                    //
                    SetCellValue(sheet, coords.EarthlyBranch.Row, coords.EarthlyBranch.Col, palace.EarthlyBranch);
                }
            }

            using var memoryStream = new MemoryStream();
            workbook.Write(memoryStream);
            return memoryStream.ToArray();
        }


        private void SetCellValue(ISheet sheet, int rowIndex, int colIndex, string value)
        {
            if (string.IsNullOrEmpty(value?.Trim())) return;
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(colIndex) ?? row.CreateCell(colIndex);
            cell.SetCellValue(value);
        }

        // 大運專用：寫入並強制黑色字體
        private void SetCellValueBlack(ISheet sheet, int rowIndex, int colIndex, string value)
        {
            if (string.IsNullOrEmpty(value?.Trim())) return;
            IRow row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            ICell cell = row.GetCell(colIndex) ?? row.CreateCell(colIndex);
            cell.SetCellValue(value);

            var workbook = sheet.Workbook;
            var origFont = cell.CellStyle.GetFont(workbook);
            IFont font = workbook.CreateFont();
            font.FontName = origFont.FontName;
            font.FontHeightInPoints = origFont.FontHeightInPoints;
            font.IsBold = origFont.IsBold;
            font.Color = NPOI.HSSF.Util.HSSFColor.Black.Index;
            ICellStyle newStyle = workbook.CreateCellStyle();
            newStyle.CloneStyleFrom(cell.CellStyle);
            newStyle.SetFont(font);
            cell.CellStyle = newStyle;
        }
    }
}