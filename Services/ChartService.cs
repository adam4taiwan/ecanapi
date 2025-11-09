// Services/ChartService.cs (最終完整版 - 包含完整十二宮與八字解析)

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using Ecanapi.Models;
using System.Text.RegularExpressions;

namespace Ecanapi.Services
{
    public class ChartService : IChartService
    {
        public async Task<ChartDataDto> ProcessChartFileAsync(IFormFile file)
        {
            var chartData = new ChartDataDto { FileName = file.FileName };
            using (var stream = new MemoryStream())
            {
                await file.CopyToAsync(stream);
                stream.Position = 0;

                IWorkbook workbook;
                string fileExtension = Path.GetExtension(file.FileName).ToLower();

                if (fileExtension == ".xls") workbook = new HSSFWorkbook(stream);
                else if (fileExtension == ".xlsx") workbook = new XSSFWorkbook(stream);
                else throw new Exception("不支援的檔案格式。");

                if (workbook.NumberOfSheets == 0) throw new Exception("Excel檔案中沒有任何工作表。");

                ISheet worksheet = workbook.GetSheetAt(0);

                // 1. 提取基本資料
                chartData.Name = worksheet.ReadCell(35, 7); // H36
                chartData.BirthDate = worksheet.ReadCell(35, 12); // M36

                // 2. 提取最完整的八字資訊
                // 年柱 (M, N 欄)
                chartData.BaziInfo.YearPillar.HeavenlyStem = worksheet.ReadCell(19, 12); // M20
                chartData.BaziInfo.YearPillar.EarthlyBranch = worksheet.ReadCell(24, 12); // M25
                chartData.BaziInfo.YearPillar.HeavenlyStemLiuShen = (worksheet.ReadCell(18, 12) + worksheet.ReadCell(18, 13)).Trim(); // M19 + N19
                chartData.BaziInfo.YearPillar.NaYin = worksheet.ReadCell(27, 12); // M28
                chartData.BaziInfo.YearPillar.HiddenStemLiuShen.AddIfNotEmpty(worksheet.ReadCell(28, 12)); chartData.BaziInfo.YearPillar.HiddenStemLiuShen.AddIfNotEmpty(worksheet.ReadCell(28, 13)); // M29, N29
                chartData.BaziInfo.YearPillar.HiddenStemLiuShen.AddIfNotEmpty(worksheet.ReadCell(29, 12)); chartData.BaziInfo.YearPillar.HiddenStemLiuShen.AddIfNotEmpty(worksheet.ReadCell(29, 13)); // M30, N30
                chartData.BaziInfo.YearPillar.HiddenStemLiuShen.AddIfNotEmpty(worksheet.ReadCell(30, 12)); chartData.BaziInfo.YearPillar.HiddenStemLiuShen.AddIfNotEmpty(worksheet.ReadCell(30, 13)); // M31, N31

                // 月柱 (K, L 欄)
                chartData.BaziInfo.MonthPillar.HeavenlyStem = worksheet.ReadCell(19, 10); // K20
                chartData.BaziInfo.MonthPillar.EarthlyBranch = worksheet.ReadCell(24, 10); // K25
                chartData.BaziInfo.MonthPillar.HeavenlyStemLiuShen = (worksheet.ReadCell(18, 10) + worksheet.ReadCell(18, 11)).Trim(); // K19 + L19
                chartData.BaziInfo.MonthPillar.NaYin = worksheet.ReadCell(27, 10); // K28
                chartData.BaziInfo.MonthPillar.HiddenStemLiuShen.AddIfNotEmpty(worksheet.ReadCell(28, 10)); chartData.BaziInfo.MonthPillar.HiddenStemLiuShen.AddIfNotEmpty(worksheet.ReadCell(28, 11)); // K29, L29
                chartData.BaziInfo.MonthPillar.HiddenStemLiuShen.AddIfNotEmpty(worksheet.ReadCell(29, 10)); chartData.BaziInfo.MonthPillar.HiddenStemLiuShen.AddIfNotEmpty(worksheet.ReadCell(29, 11)); // K30, L30
                chartData.BaziInfo.MonthPillar.HiddenStemLiuShen.AddIfNotEmpty(worksheet.ReadCell(30, 10)); chartData.BaziInfo.MonthPillar.HiddenStemLiuShen.AddIfNotEmpty(worksheet.ReadCell(30, 11)); // K31, L31

                // 日柱 (I, J 欄)
                chartData.BaziInfo.DayPillar.HeavenlyStem = worksheet.ReadCell(19, 8); // I20
                chartData.BaziInfo.DayPillar.EarthlyBranch = worksheet.ReadCell(24, 8); // I25
                chartData.BaziInfo.DayPillar.HeavenlyStemLiuShen = (worksheet.ReadCell(18, 8) + worksheet.ReadCell(18, 9)).Trim(); // I19 + J19
                chartData.BaziInfo.DayPillar.NaYin = worksheet.ReadCell(27, 8); // I28
                chartData.BaziInfo.DayPillar.HiddenStemLiuShen.AddIfNotEmpty(worksheet.ReadCell(28, 8)); chartData.BaziInfo.DayPillar.HiddenStemLiuShen.AddIfNotEmpty(worksheet.ReadCell(28, 9)); // I29, J29
                chartData.BaziInfo.DayPillar.HiddenStemLiuShen.AddIfNotEmpty(worksheet.ReadCell(29, 8)); chartData.BaziInfo.DayPillar.HiddenStemLiuShen.AddIfNotEmpty(worksheet.ReadCell(29, 9)); // I30, J30
                chartData.BaziInfo.DayPillar.HiddenStemLiuShen.AddIfNotEmpty(worksheet.ReadCell(30, 8)); chartData.BaziInfo.DayPillar.HiddenStemLiuShen.AddIfNotEmpty(worksheet.ReadCell(30, 9)); // I31, J31

                // 時柱 (G, H 欄)
                chartData.BaziInfo.TimePillar.HeavenlyStem = worksheet.ReadCell(19, 6); // G20
                chartData.BaziInfo.TimePillar.EarthlyBranch = worksheet.ReadCell(24, 6); // G25
                chartData.BaziInfo.TimePillar.HeavenlyStemLiuShen = (worksheet.ReadCell(18, 6) + worksheet.ReadCell(18, 7)).Trim(); // G19 + H19
                chartData.BaziInfo.TimePillar.NaYin = worksheet.ReadCell(27, 6); // G28
                chartData.BaziInfo.TimePillar.HiddenStemLiuShen.AddIfNotEmpty(worksheet.ReadCell(28, 6)); chartData.BaziInfo.TimePillar.HiddenStemLiuShen.AddIfNotEmpty(worksheet.ReadCell(28, 7)); // G29, H29
                chartData.BaziInfo.TimePillar.HiddenStemLiuShen.AddIfNotEmpty(worksheet.ReadCell(29, 6)); chartData.BaziInfo.TimePillar.HiddenStemLiuShen.AddIfNotEmpty(worksheet.ReadCell(29, 7)); // G30, H30
                chartData.BaziInfo.TimePillar.HiddenStemLiuShen.AddIfNotEmpty(worksheet.ReadCell(30, 6)); chartData.BaziInfo.TimePillar.HiddenStemLiuShen.AddIfNotEmpty(worksheet.ReadCell(30, 7)); // G31, H31


                // 3. 提取最完整的十二宮資訊
                var palaceCoordinates = new List<(int startRow, int startCol)>
                {
                    (6, 0), (6, 5), (6, 10), (6, 15),
                    (16, 0), (16, 15),
                    (26, 0), (26, 15),
                    (36, 0), (36, 5), (36, 10), (36, 15)
                };

                foreach (var (startRow, startCol) in palaceCoordinates)
                {
                    // === 關鍵修正：從宮位名稱的儲存格來判斷身宮 ===
                    string rawPalaceName = worksheet.ReadCell(startRow, startCol + 2); // 例如 C7, C17, C27...
                    var palace = new PalaceDto
                    {
                        // === 關鍵修正：使用新的輔助函式來正規化名稱 ===
                        Name = NormalizePalaceName(rawPalaceName),
                        IsBodyPalace = rawPalaceName.Contains("身"),
                        // ==========================================
                        //Name = palaceNameText.Replace("身", "").Trim(),
                        //IsBodyPalace = palaceNameText.Contains("身"), // 如果包含"身"字，就將 IsBodyPalace 設為 true
                        EarthlyBranch = worksheet.ReadCell(startRow + 9, startCol),
                        MainStarBrightness = worksheet.ReadCell(startRow + 2, startCol + 3),
                        PalaceAuspiciousness = worksheet.ReadCell(startRow + 2, startCol + 4)
                    };
                    // ============================
                    palace.MainStars.AddIfNotEmpty(worksheet.ReadCell(startRow, startCol + 3));
                    palace.MainStars.AddIfNotEmpty(worksheet.ReadCell(startRow, startCol + 4));
                    palace.SecondaryStars.AddIfNotEmpty(worksheet.ReadCell(startRow + 1, startCol + 1));
                    palace.SecondaryStars.AddIfNotEmpty(worksheet.ReadCell(startRow + 1, startCol + 2));
                    palace.SecondaryStars.AddIfNotEmpty(worksheet.ReadCell(startRow + 1, startCol + 3));
                    palace.SecondaryStars.AddIfNotEmpty(worksheet.ReadCell(startRow + 1, startCol + 4));
                    string minorStarsText = worksheet.ReadCell(startRow + 3, startCol);
                    if (!string.IsNullOrWhiteSpace(minorStarsText))
                    {
                        palace.MinorStars.AddRange(minorStarsText.Trim().Select(c => c.ToString()));
                    }
                    palace.AnnualStarTransformations.AddIfNotEmpty(worksheet.ReadCell(startRow + 3, startCol + 3));
                    palace.AnnualStarTransformations.AddIfNotEmpty(worksheet.ReadCell(startRow + 3, startCol + 4));
                    string a15Text = worksheet.ReadCell(startRow + 8, startCol);
                    var parts = a15Text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length > 0)
                    {
                        var stemPart = parts[0];
                        var stemSplit = stemPart.Split('-');
                        if (stemSplit.Length == 2)
                        {
                            palace.PalaceStem = stemSplit[0];
                            palace.PalaceStemTransformations = stemSplit[1];
                        }
                        palace.LifeCycleStage = parts.LastOrDefault();
                        if (parts.Length > 2)
                        {
                            palace.YearlyGeneralStars.AddRange(string.Join("", parts.Skip(1).Take(parts.Length - 2)).Select(c => c.ToString()));
                        }
                    }
                    chartData.Palaces.Add(palace);
                    // === 關鍵修正：清理並解析大運數字 ===
                    string startAgeRaw = worksheet.ReadCell(startRow, startCol);
                    string endAgeRaw = worksheet.ReadCell(startRow, startCol + 1);

                    // 使用正規表示式移除所有非數字字元
                    string startAgeClean = Regex.Replace(startAgeRaw, "[^0-9]", "");
                    string endAgeClean = Regex.Replace(endAgeRaw, "[^0-9]", "");

                    int.TryParse(startAgeClean, out int startAge);
                    int.TryParse(endAgeClean, out int endAge);

                    if (startAge > 0)
                    {
                        chartData.DecadeLuckCycles.Add(new LuckCyclePeriodDto
                        {
                            StartAge = startAge,
                            EndAge = endAge,
                            AssociatedPalace = palace.Name
                        });
                    }
 
                }

                chartData.DecadeLuckCycles = chartData.DecadeLuckCycles.OrderBy(c => c.StartAge).ToList();
                return chartData;
            }
        }
        // === 新增的輔助函式，專門處理宮位名稱 ===
        private string NormalizePalaceName(string rawName)
        {
            string cleanedName = rawName.Replace("【", "").Replace("】", "").Trim();

            if (cleanedName.Contains("身"))
            {
                // 移除 "身" 字並進行轉換
                string abbreviation = cleanedName.Replace("身", "");
                switch (abbreviation)
                {
                    case "命": return "命宮";
                    case "夫": return "夫妻";
                    case "財": return "財帛";
                    case "遷": return "遷移";
                    case "官": return "官祿";
                    case "福": return "福德";
                    default: return cleanedName; // 如果有其他情況，則回傳原始清理後的名稱
                }
            }
            // 如果不含 "身"，則直接回傳清理後的名稱
            return cleanedName;
        }
    }

    public static class NpoiExtensions
    {
        public static string ReadCell(this ISheet sheet, int row, int col)
        {
            return sheet.GetRow(row)?.GetCell(col)?.ToString()?.Trim() ?? "";
        }

        public static void AddIfNotEmpty(this List<string> list, string? value)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                list.Add(value);
            }
        }
    }
}