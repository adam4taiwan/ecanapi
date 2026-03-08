using Ecanapi.Models;
using Ecanapi.Models.Analysis;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.UserModel.Charts;
using NPOI.Util;
using NPOI.WP.UserModel;
using NPOI.XWPF.UserModel;
using System.IO;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;
using static Ecanapi.Models.Analysis.Star;
using static Ecanapi.Services.AnalysisReportService.ShiShenIntuitionEngine;

namespace Ecanapi.Services
{
    public class AnalysisReportService : IAnalysisReportService
    {
        private readonly IAnalysisService _analysisDbService;

        public AnalysisReportService(IAnalysisService analysisDbService)
        {
            _analysisDbService = analysisDbService;
        }

        public async Task<byte[]> GenerateReportAsync(AstrologyChartResult chartData, AstrologyRequest request)
        {
            // 1. 計算新的分析結果
            var analysis = CalculateGeneralStrength(chartData.Bazi);

            // 2. 產生更新後的資料副本 (這步沒錯)
            var updatedData = chartData with { BaziAnalysisResult = analysis };

            using (var memoryStream = new MemoryStream())
            {
                using (var document = new XWPFDocument())
                {
                    string baseDir = AppDomain.CurrentDomain.BaseDirectory;

                    // 思考點：在 Docker (Fly.io) 中，wwwroot 通常會被複製到執行目錄旁
                    // 使用 Path.Combine 解決 Windows (\) 與 Linux (/) 路徑斜線不同的問題
                    string bgPath = Path.Combine(baseDir, "wwwroot", "images", "中國風格書寫紙.png");
                    string sigPath = Path.Combine(baseDir, "wwwroot", "images", "signature.png");
                    string stampPath = Path.Combine(baseDir, "wwwroot", "images", "stamp.png");

                    AddCoverHeroImage(document);
                    // --- 原有的內容產生邏輯 ---
                    AddCoverPage(document, updatedData); // 封面
                    // ... 中間的其他內容 (八字、紫微、分析等) ...

                    await AddChapter1(document, updatedData, request);
                    // 最後加入簽名
                    AddSignatureAndStamp(document);
                    document.Write(memoryStream);
                }

                return memoryStream.ToArray();
            }
        }
        private void AddCoverHeroImage(XWPFDocument doc)
        {
            // 建立一個段落來放圖片
            var p = doc.CreateParagraph();
            p.Alignment = ParagraphAlignment.CENTER; // 橫向置中

            // --- 調整縱向位置：將圖片推向頁面中間 ---
            // 1440 twips 約等於 1 英吋。若要放中間，可以設大一點
            p.SpacingBefore = 4000;

            var run = p.CreateRun();

            string imgPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot", "images", "cover_page.jpg");

            if (File.Exists(imgPath))
            {
                using (var fs = new FileStream(imgPath, FileMode.Open, FileAccess.Read))
                {
                    // 這裡建議維持「橫向矩形」比例，以免圖片太高把文字擠走
                    run.AddPicture(fs, (int)NPOI.XWPF.UserModel.PictureType.JPEG, "cover", Units.ToEMU(450), Units.ToEMU(180));
                }
            }

            run.AddBreak(BreakType.PAGE);
            // 如果你希望文字接在下面，只需加幾個普通換行即可
            //run.AddBreak();
            //run.AddBreak();

        }
        //private void AddCoverHeroImage(XWPFDocument doc)
        //{
        //    var p = doc.CreateParagraph();
        //    p.Alignment = ParagraphAlignment.CENTER;
        //    var run = p.CreateRun();

        //    // 取得圖片路徑
        //    string imgPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot", "images", "cover_page.jpg");

        //    if (File.Exists(imgPath))
        //    {
        //        using (var fs = new FileStream(imgPath, FileMode.Open, FileAccess.Read))
        //        {
        //            // 寬度設為 500 EMU (約 17公分)，高度設為 200 EMU
        //            run.AddPicture(fs, (int)NPOI.XWPF.UserModel.PictureType.JPEG, "cover", Units.ToEMU(450), Units.ToEMU(180));
        //        }
        //    }
        //    run.AddBreak(); // 圖片後空一行
        //}

        private void AddCoverImage(XWPFDocument doc)
        {
            var p = doc.CreateParagraph();
            p.Alignment = ParagraphAlignment.CENTER;
            var run = p.CreateRun();

            string imgPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot", "images", "cover_hero.jpg");
            if (File.Exists(imgPath))
            {
                using (var fs = new FileStream(imgPath, FileMode.Open, FileAccess.Read))
                {
                    // 插入一張寬度適中的圖 (約 15cm)
                    run.AddPicture(fs, (int)NPOI.XWPF.UserModel.PictureType.JPEG, "cover", Units.ToEMU(450), Units.ToEMU(200));
                }
            }
            run.AddBreak(BreakType.PAGE); // 插入圖片後強制換頁，確保後續文字正確
        }
        // --- 旺衰計算核心方法 ---
        private ShiShenResult CalculateGeneralStrength(BaziInfo bazi)
        {
            var analysis = new ShiShenResult();
            string dayStem = bazi.DayPillar.HeavenlyStem;
            string monthBranch = bazi.MonthPillar.EarthlyBranch;
            double score = 0;

            string selfEl = GetElement(dayStem);
            string monthEl = GetBranchElement(monthBranch);

            // 1. 月令基本分 (50% 強度)
            if (selfEl == monthEl) score = 40; // 同氣 (旺)
            else if (IsProducedBy(monthEl, selfEl)) score = 30; // 受生 (相)
            else if (IsRestrictedBy(monthEl, selfEl)) score = -40; // 受克 (死)
            else if (IsProducedBy(selfEl, monthEl)) score = -20; // 洩氣 (休)
            else score = -10; // 耗氣 (囚)

            // 2. 地支作用力 (沖則力量減半)
            if (CheckMonthClash(bazi)) score *= 0.5;

            // 3. 通根與天干輔助
            score += CalculateRootStrength(dayStem, bazi);
            score += CalculateStemStrength(dayStem, bazi);

            analysis.Score = score;
            analysis.SeasonEffect = score < 0 ? "失令" : "得令";
            analysis.Status = score < 0 ? "身弱用印" : "正格身強";

            return analysis;
        }

        // --- 五行基礎工具 (解決名稱不存在問題) ---

        private string GetElement(string stemOrBranch)
        {
            var map = new Dictionary<string, string> {
        {"甲","木"},{"乙","木"},{"丙","火"},{"丁","火"},{"戊","土"},{"己","土"},{"庚","金"},{"辛","金"},{"壬","水"},{"癸","水"},
        {"寅","木"},{"卯","木"},{"辰","土"},{"巳","火"},{"午","火"},{"未","土"},{"申","金"},{"酉","金"},{"戌","土"},{"亥","水"},{"子","水"},{"丑","土"}
    };
            return map.ContainsKey(stemOrBranch) ? map[stemOrBranch] : "";
        }

        private string GetBranchElement(string branch) => GetElement(branch);

        private bool IsProducedBy(string source, string target)
        {
            if (source == "木" && target == "火") return true;
            if (source == "火" && target == "土") return true;
            if (source == "土" && target == "金") return true;
            if (source == "金" && target == "水") return true;
            if (source == "水" && target == "木") return true;
            return false;
        }

        private bool IsRestrictedBy(string source, string target)
        {
            if (source == "木" && target == "土") return true;
            if (source == "土" && target == "水") return true;
            if (source == "水" && target == "火") return true;
            if (source == "火" && target == "金") return true;
            if (source == "金" && target == "木") return true;
            return false;
        }

        private bool CheckMonthClash(BaziInfo bazi)
        {
            string branches = bazi.YearPillar.EarthlyBranch + bazi.DayPillar.EarthlyBranch + bazi.TimePillar.EarthlyBranch;
            var clashMap = new Dictionary<string, string> { { "子", "午" }, { "午", "子" }, { "丑", "未" }, { "未", "丑" }, { "寅", "申" }, { "申", "寅" }, { "卯", "酉" }, { "酉", "卯" }, { "辰", "戌" }, { "戌", "辰" }, { "巳", "亥" }, { "亥", "巳" } };
            return clashMap.ContainsKey(bazi.MonthPillar.EarthlyBranch) && branches.Contains(clashMap[bazi.MonthPillar.EarthlyBranch]);
        }

        private double CalculateRootStrength(string dayStem, BaziInfo bazi)
        {
            double s = 0;
            string selfEl = GetElement(dayStem);
            string allB = bazi.YearPillar.EarthlyBranch + bazi.MonthPillar.EarthlyBranch + bazi.DayPillar.EarthlyBranch + bazi.TimePillar.EarthlyBranch;
            foreach (char b in allB) { if (GetBranchElement(b.ToString()) == selfEl) s += 15; }
            return s;
        }

        private double CalculateStemStrength(string dayStem, BaziInfo bazi)
        {
            double s = 0;
            string selfEl = GetElement(dayStem);
            string allS = bazi.YearPillar.EarthlyBranch + bazi.MonthPillar.EarthlyBranch + bazi.TimePillar.EarthlyBranch;
            foreach (char st in allS) { if (GetElement(st.ToString()) == selfEl) s += 5; }
            return s;
        }
        private void AddCoverPage(XWPFDocument doc, AstrologyChartResult data)
        {
            AddParagraph(doc, $"玉洞子 先天命書", fontSize: 28, isBold: true, alignment: ParagraphAlignment.CENTER, spacingAfter: 600);
            AddParagraph(doc, "________________________", fontSize: 12, alignment: ParagraphAlignment.CENTER, spacingAfter: 600);
            var table = doc.CreateTable(3, 2);
            table.Width = 5000;
            table.GetRow(0).GetCell(0).SetText("姓名");
            table.GetRow(0).GetCell(1).SetText(data.UserName);
            table.GetRow(1).GetCell(0).SetText("陽曆生日");
            table.GetRow(1).GetCell(1).SetText(data.SolarBirthDate.ToString("yyyy年 MM月 dd日 HH時 mm分"));
            table.GetRow(2).GetCell(0).SetText("農曆生日");
            table.GetRow(2).GetCell(1).SetText(data.LunarBirthDate);
            AddParagraph(doc, "時辰恐有錯 陰騭最難憑", fontSize: 14, alignment: ParagraphAlignment.CENTER, isBold: true, color: "C00000");
            AddParagraph(doc, "萬般皆是命 半點不求人", fontSize: 14, alignment: ParagraphAlignment.CENTER, isBold: true, color: "C00000");
            doc.CreateParagraph().CreateRun().AddBreak(BreakType.PAGE);
        }
        private void SetTableToBiauKai(XWPFTable table)
        {
            if (table == null) return;

            foreach (var row in table.Rows)
            {
                foreach (var cell in row.GetTableCells())
                {
                    foreach (var para in cell.Paragraphs)
                    {
                        // 如果是 cell.SetText 產生的，確保它有 Run
                        if (para.Runs.Count == 0 && !string.IsNullOrEmpty(para.Text))
                        {
                            // 這裡不手動建立，因為 NPOI 有時會結構混亂
                        }

                        foreach (var run in para.Runs)
                        {
                            run.FontFamily = "DFKai-SB"; // 標楷體英文名
                                                         // 關鍵：必須指定 EastAsia 字體
                            run.SetFontFamily("標楷體", FontCharRange.EastAsia);

                            // 建議同時設定字級，確保統一
                            if (run.FontSize <= 0) run.FontSize = 12;
                        }
                    }
                }
            }
        }
        private async Task AddChapter1(XWPFDocument doc, AstrologyChartResult data, AstrologyRequest request)
        {
            AddParagraph(doc, "第一章：先天八字依古制定", fontSize: 20, isBold: true, color: "00008B", spacingAfter: 400);

            // 1.1 八字命盤
            AddParagraph(doc, "一、根苗花果", fontSize: 16, isBold: true, spacingAfter: 200);
            double[] skyTotal = new double[11];
            var baziTable = doc.CreateTable(6, 5);
            baziTable.Width = 5000;
            baziTable.SetColumnWidth(0, 750);
            string[] headers = { "", "時柱", "日柱", "月柱", "年柱" };
            for (int i = 0; i < headers.Length; i++)
            {
                var cell = baziTable.GetRow(0).GetCell(i);
                cell.SetText(headers[i]);
                if (cell.GetCTTc().tcPr == null) cell.GetCTTc().AddNewTcPr();
                cell.GetCTTc().tcPr.AddNewTcW().w = "5000";
            }
            baziTable.GetRow(1).GetCell(0).SetText("六神");
            baziTable.GetRow(1).GetCell(1).SetText(data.Bazi.TimePillar.HeavenlyStemLiuShen);
            baziTable.GetRow(1).GetCell(2).SetText("元神");
            baziTable.GetRow(1).GetCell(3).SetText(data.Bazi.MonthPillar.HeavenlyStemLiuShen);
            baziTable.GetRow(1).GetCell(4).SetText(data.Bazi.YearPillar.HeavenlyStemLiuShen);
            baziTable.GetRow(2).GetCell(0).SetText("天干");
            baziTable.GetRow(2).GetCell(1).SetText(data.Bazi.TimePillar.HeavenlyStem);
            baziTable.GetRow(2).GetCell(2).SetText(data.Bazi.DayPillar.HeavenlyStem);
            baziTable.GetRow(2).GetCell(3).SetText(data.Bazi.MonthPillar.HeavenlyStem);
            baziTable.GetRow(2).GetCell(4).SetText(data.Bazi.YearPillar.HeavenlyStem);
            baziTable.GetRow(3).GetCell(0).SetText("地支");
            baziTable.GetRow(3).GetCell(1).SetText(data.Bazi.TimePillar.EarthlyBranch);
            baziTable.GetRow(3).GetCell(2).SetText(data.Bazi.DayPillar.EarthlyBranch);
            baziTable.GetRow(3).GetCell(3).SetText(data.Bazi.MonthPillar.EarthlyBranch);
            baziTable.GetRow(3).GetCell(4).SetText(data.Bazi.YearPillar.EarthlyBranch);
            // 計算五行
            switch (data.Bazi.TimePillar.HeavenlyStem)
            {
                case "甲":
                    skyTotal[1] = skyTotal[1] + 1;
                    break;
                case "乙":
                    skyTotal[2] = skyTotal[2] + 1;
                    break;
                case "丙":
                    skyTotal[3] = skyTotal[3] + 1;
                    break;
                case "丁":
                    skyTotal[4] = skyTotal[4] + 1;
                    break;
                case "戊":
                    skyTotal[5] = skyTotal[5] + 1;
                    break;
                case "己":
                    skyTotal[6] = skyTotal[6] + 1;
                    break;
                case "庚":
                    skyTotal[7] = skyTotal[7] + 1;
                    break;
                case "辛":
                    skyTotal[8] = skyTotal[8] + 1;
                    break;
                case "壬":
                    skyTotal[9] = skyTotal[9] + 1;
                    break;
                case "癸":
                    skyTotal[10] = skyTotal[10] + 1;
                    break;
            }
            switch (data.Bazi.DayPillar.HeavenlyStem)
            {
                case "甲":
                    skyTotal[1] = skyTotal[1] + 1;
                    break;
                case "乙":
                    skyTotal[2] = skyTotal[2] + 1;
                    break;
                case "丙":
                    skyTotal[3] = skyTotal[3] + 1;
                    break;
                case "丁":
                    skyTotal[4] = skyTotal[4] + 1;
                    break;
                case "戊":
                    skyTotal[5] = skyTotal[5] + 1;
                    break;
                case "己":
                    skyTotal[6] = skyTotal[6] + 1;
                    break;
                case "庚":
                    skyTotal[7] = skyTotal[7] + 1;
                    break;
                case "辛":
                    skyTotal[8] = skyTotal[8] + 1;
                    break;
                case "壬":
                    skyTotal[9] = skyTotal[9] + 1;
                    break;
                case "癸":
                    skyTotal[10] = skyTotal[10] + 1;
                    break;
            }

            switch (data.Bazi.MonthPillar.HeavenlyStem)
            {
                case "甲":
                    skyTotal[1] = skyTotal[1] + 1;
                    break;
                case "乙":
                    skyTotal[2] = skyTotal[2] + 1;
                    break;
                case "丙":
                    skyTotal[3] = skyTotal[3] + 1;
                    break;
                case "丁":
                    skyTotal[4] = skyTotal[4] + 1;
                    break;
                case "戊":
                    skyTotal[5] = skyTotal[5] + 1;
                    break;
                case "己":
                    skyTotal[6] = skyTotal[6] + 1;
                    break;
                case "庚":
                    skyTotal[7] = skyTotal[7] + 1;
                    break;
                case "辛":
                    skyTotal[8] = skyTotal[8] + 1;
                    break;
                case "壬":
                    skyTotal[9] = skyTotal[9] + 1;
                    break;
                case "癸":
                    skyTotal[10] = skyTotal[10] + 1;
                    break;
            }

            switch (data.Bazi.YearPillar.HeavenlyStem)
            {
                case "甲":
                    skyTotal[1] = skyTotal[1] + 1;
                    break;
                case "乙":
                    skyTotal[2] = skyTotal[2] + 1;
                    break;
                case "丙":
                    skyTotal[3] = skyTotal[3] + 1;
                    break;
                case "丁":
                    skyTotal[4] = skyTotal[4] + 1;
                    break;
                case "戊":
                    skyTotal[5] = skyTotal[5] + 1;
                    break;
                case "己":
                    skyTotal[6] = skyTotal[6] + 1;
                    break;
                case "庚":
                    skyTotal[7] = skyTotal[7] + 1;
                    break;
                case "辛":
                    skyTotal[8] = skyTotal[8] + 1;
                    break;
                case "壬":
                    skyTotal[9] = skyTotal[9] + 1;
                    break;
                case "癸":
                    skyTotal[10] = skyTotal[10] + 1;
                    break;
            }
            //藏干六神
            var strYear = "";
            for (int i = 0; i < data.Bazi.YearPillar.HiddenStemLiuShen.Count; i++)
            {
                if ((i & 1) == 1)
                {
                    switch (data.Bazi.YearPillar.HiddenStemLiuShen[i])
                    {
                        case "甲":
                            skyTotal[1] = skyTotal[1] + 0.5;
                            break;
                        case "乙":
                            skyTotal[2] = skyTotal[2] + 1;
                            break;
                        case "丙":
                            skyTotal[3] = skyTotal[3] + 0.5;
                            break;
                        case "丁":
                            skyTotal[4] = skyTotal[4] + 0.5;
                            break;
                        case "戊":
                            skyTotal[5] = skyTotal[5] + 0.5;
                            break;
                        case "己":
                            skyTotal[6] = skyTotal[6] + 0.5;
                            break;
                        case "庚":
                            skyTotal[7] = skyTotal[7] + 0.5;
                            break;
                        case "辛":
                            skyTotal[8] = skyTotal[8] + 1;
                            break;
                        case "壬":
                            skyTotal[9] = skyTotal[9] + 0.5;
                            break;
                        case "癸":
                            skyTotal[10] = skyTotal[10] + 1;
                            break;
                    }
                }
                strYear = strYear + data.Bazi.YearPillar.HiddenStemLiuShen[i];
            }
            var strMonth = "";
            for (int i = 0; i < data.Bazi.MonthPillar.HiddenStemLiuShen.Count; i++)
            {
                if ((i & 1) == 1)
                {
                    switch (data.Bazi.MonthPillar.HiddenStemLiuShen[i])
                    {
                        case "甲":
                            skyTotal[1] = skyTotal[1] + 1;
                            break;
                        case "乙":
                            skyTotal[2] = skyTotal[2] + 1;
                            break;
                        case "丙":
                            skyTotal[3] = skyTotal[3] + 1;
                            break;
                        case "丁":
                            skyTotal[4] = skyTotal[4] + 1;
                            break;
                        case "戊":
                            skyTotal[5] = skyTotal[5] + 1;
                            break;
                        case "己":
                            skyTotal[6] = skyTotal[6] + 1;
                            break;
                        case "庚":
                            skyTotal[7] = skyTotal[7] + 1;
                            break;
                        case "辛":
                            skyTotal[8] = skyTotal[8] + 1;
                            break;
                        case "壬":
                            skyTotal[9] = skyTotal[9] + 1;
                            break;
                        case "癸":
                            skyTotal[10] = skyTotal[10] + 1;
                            break;
                    }
                }
                strMonth = strMonth + data.Bazi.MonthPillar.HiddenStemLiuShen[i];
            }
            var strDay = "";
            for (int i = 0; i < data.Bazi.DayPillar.HiddenStemLiuShen.Count; i++)
            {
                if ((i & 1) == 1)
                {
                    switch (data.Bazi.DayPillar.HiddenStemLiuShen[i])
                    {
                        case "甲":
                            skyTotal[1] = skyTotal[1] + 0.5;
                            break;
                        case "乙":
                            skyTotal[2] = skyTotal[2] + 1;
                            break;
                        case "丙":
                            skyTotal[3] = skyTotal[3] + 0.5;
                            break;
                        case "丁":
                            skyTotal[4] = skyTotal[4] + 0.5;
                            break;
                        case "戊":
                            skyTotal[5] = skyTotal[5] + 0.5;
                            break;
                        case "己":
                            skyTotal[6] = skyTotal[6] + 0.5;
                            break;
                        case "庚":
                            skyTotal[7] = skyTotal[7] + 0.5;
                            break;
                        case "辛":
                            skyTotal[8] = skyTotal[8] + 1;
                            break;
                        case "壬":
                            skyTotal[9] = skyTotal[9] + 0.5;
                            break;
                        case "癸":
                            skyTotal[10] = skyTotal[10] + 1;
                            break;
                    }
                }
                strDay = strDay + data.Bazi.DayPillar.HiddenStemLiuShen[i];
            }
            var strTime = "";
            for (int i = 0; i < data.Bazi.TimePillar.HiddenStemLiuShen.Count; i++)
            {
                if ((i & 1) == 1)
                {
                    switch (data.Bazi.TimePillar.HiddenStemLiuShen[i])
                    {
                        case "甲":
                            skyTotal[1] = skyTotal[1] + 0.5;
                            break;
                        case "乙":
                            skyTotal[2] = skyTotal[2] + 1;
                            break;
                        case "丙":
                            skyTotal[3] = skyTotal[3] + 0.5;
                            break;
                        case "丁":
                            skyTotal[4] = skyTotal[4] + 0.5;
                            break;
                        case "戊":
                            skyTotal[5] = skyTotal[5] + 0.5;
                            break;
                        case "己":
                            skyTotal[6] = skyTotal[6] + 0.5;
                            break;
                        case "庚":
                            skyTotal[7] = skyTotal[7] + 0.5;
                            break;
                        case "辛":
                            skyTotal[8] = skyTotal[8] + 1;
                            break;
                        case "壬":
                            skyTotal[9] = skyTotal[9] + 0.5;
                            break;
                        case "癸":
                            skyTotal[10] = skyTotal[10] + 1;
                            break;
                    }
                }
                strTime = strTime+ data.Bazi.TimePillar.HiddenStemLiuShen[i];
            }
            baziTable.GetRow(4).GetCell(0).SetText("藏神");
            baziTable.GetRow(4).GetCell(1).SetText(strTime);
            baziTable.GetRow(4).GetCell(2).SetText(strDay);
            baziTable.GetRow(4).GetCell(3).SetText(strMonth);
            baziTable.GetRow(4).GetCell(4).SetText(strYear);
            baziTable.GetRow(5).GetCell(0).SetText("納音");
            baziTable.GetRow(5).GetCell(1).SetText(data.Bazi.TimePillar.NaYin);
            baziTable.GetRow(5).GetCell(2).SetText(data.Bazi.DayPillar.NaYin);
            baziTable.GetRow(5).GetCell(3).SetText(data.Bazi.MonthPillar.NaYin);
            baziTable.GetRow(5).GetCell(4).SetText(data.Bazi.YearPillar.NaYin);
            SetTableToBiauKai(baziTable);
            AddParagraph(doc, "");
            var table1 = doc.CreateTable(2, 6); // +2 為標題與總結
            table1.Width = 5000;
            table1.GetRow(0).GetCell(0).SetText("");
            var Col1 = table1.GetRow(0).GetCell(0).GetCTTc().AddNewTcPr();
            Col1.tcW = new CT_TblWidth { w = "300", type = ST_TblWidth.dxa };
            table1.GetRow(0).GetCell(1).SetText("旺");
            table1.GetRow(0).GetCell(2).SetText("相");
            table1.GetRow(0).GetCell(3).SetText("死");
            table1.GetRow(0).GetCell(4).SetText("休");
            table1.GetRow(0).GetCell(5).SetText("囚");
            switch (data.Bazi.MonthPillar.EarthlyBranch)
            {
                case "子":
                    table1.GetRow(1).GetCell(1).SetText("水");
                    table1.GetRow(1).GetCell(2).SetText("木");
                    table1.GetRow(1).GetCell(3).SetText("火");
                    table1.GetRow(1).GetCell(4).SetText("金");
                    table1.GetRow(1).GetCell(5).SetText("土");
                    break;
                case "丑":
                    table1.GetRow(1).GetCell(1).SetText("土");
                    table1.GetRow(1).GetCell(2).SetText("金");
                    table1.GetRow(1).GetCell(3).SetText("水");
                    table1.GetRow(1).GetCell(4).SetText("火");
                    table1.GetRow(1).GetCell(5).SetText("木");
                    break;
                case "寅":
                    table1.GetRow(1).GetCell(1).SetText("木");
                    table1.GetRow(1).GetCell(2).SetText("火");
                    table1.GetRow(1).GetCell(3).SetText("土");
                    table1.GetRow(1).GetCell(4).SetText("水");
                    table1.GetRow(1).GetCell(5).SetText("金");
                    break;
                case "卯":
                    table1.GetRow(1).GetCell(1).SetText("木");
                    table1.GetRow(1).GetCell(2).SetText("火");
                    table1.GetRow(1).GetCell(3).SetText("土");
                    table1.GetRow(1).GetCell(4).SetText("水");
                    table1.GetRow(1).GetCell(5).SetText("金");
                    break;
                case "辰":
                    table1.GetRow(1).GetCell(1).SetText("土");
                    table1.GetRow(1).GetCell(2).SetText("金");
                    table1.GetRow(1).GetCell(3).SetText("水");
                    table1.GetRow(1).GetCell(4).SetText("火");
                    table1.GetRow(1).GetCell(5).SetText("木");
                    break;
                case "巳":
                    table1.GetRow(1).GetCell(1).SetText("火");
                    table1.GetRow(1).GetCell(2).SetText("土");
                    table1.GetRow(1).GetCell(3).SetText("金");
                    table1.GetRow(1).GetCell(4).SetText("木");
                    table1.GetRow(1).GetCell(5).SetText("水");
                    break;
                case "午":
                    table1.GetRow(1).GetCell(1).SetText("火");
                    table1.GetRow(1).GetCell(2).SetText("土");
                    table1.GetRow(1).GetCell(3).SetText("金");
                    table1.GetRow(1).GetCell(4).SetText("木");
                    table1.GetRow(1).GetCell(5).SetText("水");
                    break;
                case "未":
                    table1.GetRow(1).GetCell(1).SetText("土");
                    table1.GetRow(1).GetCell(2).SetText("金");
                    table1.GetRow(1).GetCell(3).SetText("水");
                    table1.GetRow(1).GetCell(4).SetText("火");
                    table1.GetRow(1).GetCell(5).SetText("木");
                    break;
                case "申":
                    table1.GetRow(1).GetCell(1).SetText("金");
                    table1.GetRow(1).GetCell(2).SetText("水");
                    table1.GetRow(1).GetCell(3).SetText("木");
                    table1.GetRow(1).GetCell(4).SetText("土");
                    table1.GetRow(1).GetCell(5).SetText("火");
                    break;
                case "酉":
                    table1.GetRow(1).GetCell(1).SetText("金");
                    table1.GetRow(1).GetCell(2).SetText("水");
                    table1.GetRow(1).GetCell(3).SetText("木");
                    table1.GetRow(1).GetCell(4).SetText("土");
                    table1.GetRow(1).GetCell(5).SetText("火");
                    break;
                case "戌":
                    table1.GetRow(1).GetCell(1).SetText("土");
                    table1.GetRow(1).GetCell(2).SetText("金");
                    table1.GetRow(1).GetCell(3).SetText("水");
                    table1.GetRow(1).GetCell(4).SetText("火");
                    table1.GetRow(1).GetCell(5).SetText("木");
                    break;
                case "亥":
                    table1.GetRow(1).GetCell(1).SetText("水");
                    table1.GetRow(1).GetCell(2).SetText("木");
                    table1.GetRow(1).GetCell(3).SetText("火");
                    table1.GetRow(1).GetCell(4).SetText("金");
                    table1.GetRow(1).GetCell(5).SetText("土");
                    break;
            }
            SetTableToBiauKai(table1);
            // 新增干支五行生剋
            // 1.4 干支五行生剋星剎分析
            AddParagraph(doc, "干支五行生剋星剎分析", fontSize: 16, isBold: true, spacingAfter: 200);
            var stems = new[] { data.Bazi.YearPillar.HeavenlyStem, data.Bazi.MonthPillar.HeavenlyStem,
                       data.Bazi.DayPillar.HeavenlyStem, data.Bazi.TimePillar.HeavenlyStem };
            var branches = new[] { data.Bazi.YearPillar.EarthlyBranch, data.Bazi.MonthPillar.EarthlyBranch,
                          data.Bazi.DayPillar.EarthlyBranch, data.Bazi.TimePillar.EarthlyBranch };
            var stemCount = stems.GroupBy(s => s).ToDictionary(g => g.Key, g => g.Count());
            var branchCount = branches.GroupBy(b => b).ToDictionary(g => g.Key, g => g.Count());

            var fiveElements = new Dictionary<string, string>
            {
                { "甲", "木" }, { "乙", "木" }, { "丙", "火" }, { "丁", "火" }, { "戊", "土" },
                { "己", "土" }, { "庚", "金" }, { "辛", "金" }, { "壬", "水" }, { "癸", "水" },
                { "子", "水" }, { "丑", "土" }, { "寅", "木" }, { "卯", "木" }, { "辰", "土" },
                { "巳", "火" }, { "午", "火" }, { "未", "土" }, { "申", "金" }, { "酉", "金" },
                { "戌", "土" }, { "亥", "水" }
            };
            var relationships = new Dictionary<(string, string), string>
            {
                { ("子", "午"), "沖" }, { ("丑", "未"), "沖" }, { ("寅", "申"), "沖" },
                { ("卯", "酉"), "沖" }, { ("辰", "戌"), "沖" }, { ("巳", "亥"), "沖" },
                { ("子", "丑"), "合" }, { ("寅", "亥"), "合" }, { ("卯", "戌"), "合" },
                { ("巳", "申"), "合" }, { ("午", "未"), "合" }, { ("辰", "酉"), "合" },
                { ("寅", "巳"), "害" }, { ("丑", "午"), "害" }, { ("戌", "酉"), "害" },
                { ("卯", "辰"), "害" }, { ("申", "亥"), "害" }, { ("子", "未"), "害" },
                { ("子", "酉"), "破" }, { ("卯", "午"), "破" }, { ("丑", "辰"), "破" }, { ("未", "戌"), "破" }
            };
            var threeCriminals = new Dictionary<string[], string>
            {
                { new[] { "寅", "巳", "申" }, "三刑" },
                { new[] { "丑", "未", "戌" }, "三刑" }
            };
            var table = doc.CreateTable(11, 16); // +2 為標題與總結
            table.Width = 5000;

            // Column 2 auto resizes

            // Set last column width
            //var mPrLast = table.GetRow(0).GetCell(2).GetCTTc().AddNewTcPr();
            //mPrLast.tcW = new CT_TblWidth { w = "700", type = ST_TblWidth.dxa };


            table.GetRow(0).GetCell(0).SetText("");
            var Col0 = table.GetRow(0).GetCell(0).GetCTTc().AddNewTcPr();
            Col0.tcW = new CT_TblWidth { w = "300", type = ST_TblWidth.dxa };
            table.GetRow(0).GetCell(1).SetText("干");
            table.GetRow(0).GetCell(2).SetText("甲");
            table.GetRow(0).GetCell(3).SetText("乙");
            table.GetRow(0).GetCell(4).SetText("丙");
            table.GetRow(0).GetCell(5).SetText("丁");
            table.GetRow(0).GetCell(6).SetText("戊");
            table.GetRow(0).GetCell(7).SetText("己");
            table.GetRow(0).GetCell(8).SetText("庚");
            table.GetRow(0).GetCell(9).SetText("辛");
            table.GetRow(0).GetCell(10).SetText("壬");
            table.GetRow(0).GetCell(11).SetText("癸");
            table.GetRow(0).GetCell(12).SetText("");
            table.GetRow(0).GetCell(13).SetText("");
            table.GetRow(0).GetCell(14).SetText("五行");
            table.GetRow(0).GetCell(15).SetText("星剎");
            ////-----------------------------------------
            int rowIdx = 1;
            foreach (var stem in stemCount.Keys.Where(s => s != null))
            {
                switch (rowIdx)
                {
                    case 1: table.GetRow(rowIdx).GetCell(0).SetText(data.Bazi.YearPillar.HeavenlyStemLiuShen); break;
                    case 2: table.GetRow(rowIdx).GetCell(0).SetText(data.Bazi.MonthPillar.HeavenlyStemLiuShen); break;
                    case 3: table.GetRow(rowIdx).GetCell(0).SetText("日"); break;
                    case 4: table.GetRow(rowIdx).GetCell(0).SetText(data.Bazi.TimePillar.HeavenlyStemLiuShen); break;
                }
                //table.GetRow(rowIdx).GetCell(0).SetText("干");
                table.GetRow(rowIdx).GetCell(1).SetText(stem);
                // 年 SELECT * FROM public."干組" where "CU_NO" = '辛'
                var sQuery = "SELECT CONCAT(CONCAT(\"CU_1\", '甲'), CONCAT(\"CU_2\", '乙') , CONCAT(\"CU_3\", '丙') , CONCAT(\"CU_4\", '丁') , CONCAT(\"CU_5\", '戊') , CONCAT(\"CU_6\", '己') , CONCAT(\"CU_7\", '庚') ,CONCAT(\"CU_8\", '辛') ,CONCAT(\"CU_9\", '壬') , CONCAT(\"CU_10\", '癸') ) as \"Value\" FROM public.\"干六神\"  where \"CU_NO\" = ";
                sQuery = sQuery + " '" + stem + "' ";
                var resList = await _analysisDbService.ExecuteRawQueryAsync(sQuery);
                for (int i = 0; i < 10; i++)
                {
                    table.GetRow(rowIdx).GetCell(i + 2).SetText(resList.Substring(i*2,2));
                }
                //table.GetRow(rowIdx).GetCell(3).SetText(stemCount[stem].ToString());
                //var element = fiveElements.ContainsKey(stem) ? fiveElements[stem] : "未知";
                //var relations = branches.Select(b => (stem, b))
                //    .Select(t => relationships.FirstOrDefault(r => r.Key.Item1 == t.Item1 && r.Key.Item2 == t.Item2).Value ?? "")
                //    .Where(r => !string.IsNullOrEmpty(r))
                //    .Distinct();
                //// 檢查三刑
                //var threeCrimRelation = threeCriminals.FirstOrDefault(tc => tc.Key.Contains(stem) && branches.Intersect(tc.Key).Count() >= 2).Value;
                //if (!string.IsNullOrEmpty(threeCrimRelation)) relations = relations.Append(threeCrimRelation).Distinct();
                //table.GetRow(rowIdx).GetCell(4).SetText($"{element} ({string.Join(", ", relations)})");
                rowIdx++;
            }
            table.GetRow(5).GetCell(0).SetText("");
            Col0 = table.GetRow(5).GetCell(0).GetCTTc().AddNewTcPr();
            Col0.tcW = new CT_TblWidth { w = "300", type = ST_TblWidth.dxa };
            table.GetRow(5).GetCell(1).SetText("支");
            table.GetRow(5).GetCell(2).SetText("子");
            table.GetRow(5).GetCell(3).SetText("丑");
            table.GetRow(5).GetCell(4).SetText("寅");
            table.GetRow(5).GetCell(5).SetText("卯");
            table.GetRow(5).GetCell(6).SetText("辰");
            table.GetRow(5).GetCell(7).SetText("巳");
            table.GetRow(5).GetCell(8).SetText("午");
            table.GetRow(5).GetCell(9).SetText("未");
            table.GetRow(5).GetCell(10).SetText("申");
            table.GetRow(5).GetCell(11).SetText("酉");
            table.GetRow(5).GetCell(12).SetText("戌");
            table.GetRow(5).GetCell(13).SetText("亥");
            table.GetRow(5).GetCell(14).SetText("五行");
            table.GetRow(5).GetCell(15).SetText("星剎");
            rowIdx=6;
            foreach (var branch in branchCount.Keys.Where(b => b != null))
            {
                switch (rowIdx)
                {
                    case 6: table.GetRow(rowIdx).GetCell(0).SetText(data.Bazi.YearPillar.HiddenStemLiuShen[0]); break;
                    case 7: table.GetRow(rowIdx).GetCell(0).SetText(data.Bazi.MonthPillar.HiddenStemLiuShen[0]); break;
                    case 8: table.GetRow(rowIdx).GetCell(0).SetText(data.Bazi.DayPillar.HiddenStemLiuShen[0]); break;
                    case 9: table.GetRow(rowIdx).GetCell(0).SetText(data.Bazi.TimePillar.HiddenStemLiuShen[0]); break;
                }
                //table.GetRow(rowIdx).GetCell(0).SetText("支");
                table.GetRow(rowIdx).GetCell(1).SetText(branch);
                //
                // 年 SELECT * FROM public."干組" where "CU_NO" = '辛'
                var sQuery = "SELECT CONCAT(CONCAT(\"CU_1\", '子'), CONCAT(\"CU_2\", '丑'), CONCAT(\"CU_3\", '寅'), CONCAT(\"CU_4\", '卯') , CONCAT(\"CU_5\", '辰') , CONCAT(\"CU_6\", '巳') , CONCAT(\"CU_7\", '午') ,CONCAT(\"CU_8\", '未') ,CONCAT(\"CU_9\", '申') , CONCAT(\"CU_10\", '酉') , CONCAT(\"CU_11\", '戌') , CONCAT(\"CU_12\", '亥')) as \"Value\" FROM public.\"支組\" where \"CU_NO\" =";
                sQuery = sQuery + " '" + branch + "' ";
                var resList = await _analysisDbService.ExecuteRawQueryAsync(sQuery);
                for (int i = 0; i < 12; i++)
                {
                    table.GetRow(rowIdx).GetCell(i + 2).SetText(resList.Substring(i * 2, 2));
                    //table.GetRow(rowIdx).GetCell(i + 2).SetText(resList[i].ToString());
                }
                //table.GetRow(rowIdx).GetCell(3).SetText(branchCount[branch].ToString());
                //var element = fiveElements.ContainsKey(branch) ? fiveElements[branch] : "未知";
                //var relations = stems.Select(s => (s, branch))
                //    .Select(t => relationships.FirstOrDefault(r => r.Key.Item1 == t.Item1 && r.Key.Item2 == t.Item2).Value ?? "")
                //    .Where(r => !string.IsNullOrEmpty(r))
                //    .Distinct();
                //// 檢查三刑
                //var threeCrimRelation = threeCriminals.FirstOrDefault(tc => tc.Key.Contains(branch) && branches.Intersect(tc.Key).Count() >= 2).Value;
                //if (!string.IsNullOrEmpty(threeCrimRelation)) relations = relations.Append(threeCrimRelation).Distinct();
                //table.GetRow(rowIdx).GetCell(4).SetText($"{element} ({string.Join(", ", relations)})");
                rowIdx++;
            }
            table.GetRow(10).GetCell(1).SetText("計");
            table.GetRow(10).GetCell(2).SetText(skyTotal[1].ToString());
            table.GetRow(10).GetCell(3).SetText(skyTotal[2].ToString());
            table.GetRow(10).GetCell(4).SetText(skyTotal[3].ToString());
            table.GetRow(10).GetCell(5).SetText(skyTotal[4].ToString());
            table.GetRow(10).GetCell(6).SetText(skyTotal[5].ToString());
            table.GetRow(10).GetCell(7).SetText(skyTotal[6].ToString());
            table.GetRow(10).GetCell(8).SetText(skyTotal[7].ToString());
            table.GetRow(10).GetCell(9).SetText(skyTotal[8].ToString());
            table.GetRow(10).GetCell(10).SetText(skyTotal[9].ToString());
            table.GetRow(10).GetCell(11).SetText(skyTotal[10].ToString());
            SetTableToBiauKai(table);
            // 總結
            //var allRelations = stems.SelectMany(s => branches.Select(b => (s, b)))
            //    .Select(t => relationships.FirstOrDefault(r => r.Key.Item1 == t.Item1 && r.Key.Item2 == t.Item2).Value ?? "")
            //    .Where(r => !string.IsNullOrEmpty(r))
            //    .Distinct()
            //    .ToList();
            //// 加入三刑總結
            //var allThreeCrim = threeCriminals.Where(tc => branches.Intersect(tc.Key).Count() == 3).Select(tc => tc.Value).Distinct();
            //allRelations.AddRange(allThreeCrim);
            //table.GetRow(rowIdx).GetCell(0).SetText("總結");
            //table.GetRow(rowIdx).GetCell(1).SetText("-");
            //table.GetRow(rowIdx).GetCell(2).SetText("-");
            //table.GetRow(rowIdx).GetCell(3).SetText("-");
            //table.GetRow(rowIdx).GetCell(4).SetText($"整體星剎關係：{string.Join(", ", allRelations)}");
            // 結束表格
            // SELECT CONCAT(rgcz,xgfx,aqfx,syfx,cyfx,jkfx) FROM public.六十甲子命主
            //var strQuery = "SELECT CONCAT(rgcz,xgfx,aqfx,syfx,cyfx,jkfx) AS \"Value\" FROM public.六十甲子命主" + " where rgz = ";
            var strQuery = "SELECT CONCAT(rgcz) AS \"Value\" FROM public.六十甲子命主" + " where rgz = ";
            strQuery = strQuery + " '"+data.Bazi.DayPillar.HeavenlyStem + data.Bazi.DayPillar.EarthlyBranch+"' ";
            var resultsql = await _analysisDbService.ExecuteRawQueryAsync(strQuery);
            AddParagraph(doc, " 日元分析:", fontSize: 12, isBold: true, spacingAfter: 200);
            if (resultsql != null)
            {
                AddParagraph(doc, $" {resultsql}");
            }
            AddParagraph(doc, "");

            // SELECT whw  FROM public."五行" where wh = '火'
            var     strFiveKind = "";
            switch (data.Bazi.DayMaster)
            {
                     case "甲":strFiveKind = "木" ; break;
                     case "乙":strFiveKind = "木"; break;
                     case "丙":strFiveKind = "火"; break;
                    case "丁":strFiveKind = "火"; break;  
                    case "戊":strFiveKind = "土"; break;
                    case "己":strFiveKind = "土"; break;
                    case "庚":strFiveKind = "金"; break;
                    case "辛":strFiveKind = "金"; break;
                    case "壬":strFiveKind = "水"; break;
                    case "癸":strFiveKind = "水"; break;
            }
            //strQuery = "SELECT whw as \"Value\" FROM public.五行 " + " where wh = ";
            //strQuery = strQuery + " '" + strFiveKind + "' ";
            //resultsql = await _analysisDbService.ExecuteRawQueryAsync(strQuery);
            //AddParagraph(doc, " 五行分析:", fontSize: 12, isBold: true, spacingAfter: 200);
            //if (resultsql != null)
            //{
            //    AddParagraph(doc, $" {resultsql}");
            //}
            if (strFiveKind == "木")
            {
                resultsql = @"　　　　　　　木生春月旺、水多主災殃、不是貧困守、
			    就是少年亡、最喜戊己土、見火大吉昌、
			    見官反無官、庚辛命不強、夏木不為凋、
			    火炎怕土燥、火盛要水養、水盛木則漂、
			    最喜見寅卯、子辰亦為好、財官若根輕、
			    午未運不高、木生四季節、怎把戊己克、
			    求逢壬癸水、反宜金疊疊、辰丑須用火、
			    未戍水木扶、全是火土現、土運亦可安、
			    秋木逢旺金、金旺及傷身、不宜重見土、
			    幹頭愛丙丁、金木若相戰、財運反遭陷、
			    壬癸若相助、定然名利亨、冬木枯葉調、
			    水冷怕金削、有土相逢火、溫暖發根苗、
			    冬天水木泛、名利總虛無、運逢火土助、
			    名利才自如。";
                AddParagraph(doc, $" {resultsql}");
            }
            if (strFiveKind == "火")
            {
                resultsql = @"　　　　　　　火到春見陽、木盛火必強、水旺金不旺、
			    富貴理應當、火多木必焚、最喜見官鄉、
			    支中見多卯、反倒為災殃、
			    夏火非尋常、木多主不祥、水旺要金相、
			    方才顯高強、火盛土必燥、無水焉能好、
			    宜壬不宜癸、不須寅卯繞、
			    火生四季衰、水多常有害、用金不和諧、
			    甲己喜相愛、未戍多孤寡、丑辰是良才、
			    休逢戊己旺、必定有禍殃、
			    秋月火不強、唯喜甲丙幫、用木還用土、
			    土重則為殃、歲運逢火地、自然財興旺、
			    水盛最無情、不夭眼目盲、
			    火生冬不旺、四柱怕金鄉、愛用甲乙木、
			    幹適戊己鄉、人生若逢此、必定是官郎、
			    只恐財殺旺、刑沖見閻王、";
                AddParagraph(doc, $" {resultsql}");
            }
            if (strFiveKind == "土")
            {
                resultsql = @"　　　　　　　春月生戊己、難抵甲和乙、用金必有利、
			    有火方是奇、不愁春水旺、身衰反無依、
			    水火兩相濟、方能有名利、
			    戊己夏月生、幹頭喜見金、旺遇火丙丁、
			    無水不顯榮、水火名既濟、戊己得安生、
			    但愁木火吐、焦躁身不寧、
			    四季戊己旺、不宜火焰鄉、甲乙喜相逢、
			    用水定吉昌、丑辰單用火、卯旺名聲揚、
			    戍未水木鄉、戊己忌再幫、
			    秋月逢戊己、金旺瀉土氣、用火方成器、
			    用水不相宜、此時愁木旺、怎能無災殃、
			    或是遭人欺、或是有殘疾、
			    戊己逢冬生、見金則不榮、甲乙雖為病、
			    有火是公卿、支中若無根、為人定漂流、
			    須得戊己幫、亥子是財鄉、";
                AddParagraph(doc, $" {resultsql}");
            }
            if (strFiveKind == "金")
            {
                resultsql = @"　　　　　　　金生春月地、焉能用甲乙、用土還傷火、
			    無傷不為奇、金木若相戰、兩丁名利齊、
			    若是喜水旺、休囚苦無依、
			    庚辛夏月生、每朝遇丙丁、休用甲乙木、
			    壬癸水擔承、不宜再逢火、火旺金必溶、
			    戊己土來幫、必定名利亨、
			    四季旺生金、母旺子相生、用木正用水、
			    單怕水埋金、戍未逢旺火、見水得安生、
			    丑辰用丙丁、木鄉有名聲、
			    秋月金旺重、水土兩無功、問君何所用、
			    最喜財官逢、旺火鑄金鐘、歲逢名利通、
			    英惑無根氣、貧賤九流同、
			    庚辛產於冬、冬水旺無窮、水旺金必沉、
			    早亡少壯童、支安根祿重、印生始太平、
			    雖然喜逢土、無火不顯榮、";
                AddParagraph(doc, $" {resultsql}");
            }
            if (strFiveKind == "水")
            {
                resultsql = @"　　　　　　　春水不當權、木盛奪水源、用金禮相見、
			    見火樂封源、水木傷官格、喜見財與官、
			    水木兩相見、空來人世間、
			    水生夏月天、木多火炎炎、休言財官旺、
			    只要不枯源、金水若相涵、名利兩相歡、
			    土燥金不見、不能成方圓、
			    水生四季弱、土旺不相合、用金愁逢火、
			    用木怕金多、土旺惡煞侵、一命見閻羅、
			    丑辰是其根、戍未最怕火、
			    壬癸生於秋、生人樂悠悠、戊己喜相逢、
			    有火樂風流、不宜金再旺、名高反不高、
			    財官若相旺、福壽度春秋、
			    冬水旺無窮、金木兩無功、幹透戊己土、
			    富貴丙丁逢、但愁金水旺、父星反遭沖、
			    祖業難有靠、紮根火土中、";
                AddParagraph(doc, $" {resultsql}");
            }
            AddParagraph(doc, "");
            //跳頁
            doc.CreateParagraph().CreateRun().AddBreak(BreakType.PAGE);
            // --- 六十甲子日對時 ---
            strQuery = "SELECT \"desc\" as \"Value\" FROM public.\"六十甲子日對時\" " + " where \"Sky\" =  ";
            strQuery = strQuery + " '" + data.Bazi.DayPillar.HeavenlyStem + "' and \"Month\" =  ";
            strQuery = strQuery + " '" + data.Bazi.MonthPillar.EarthlyBranch + "月' and \"time\" like";
            strQuery = strQuery + " '%" + data.Bazi.TimePillar.EarthlyBranch + "%' ";
            resultsql = await _analysisDbService.ExecuteRawQueryAsync(strQuery);
            //Task<IEnumerable<SixtyJiaziDayToHour>> GetSixtyJiaziDayToHourAsync(string? sky, string? month, string? time);
            //var resultMain = await _analysisDbService.GetSixtyJiaziDayToHourAsync(data.Bazi.DayPillar.HeavenlyStem, data.Bazi.MonthPillar.EarthlyBranch, data.Bazi.TimePillar.EarthlyBranch);
            //var pt = resultMain.FirstOrDefault();
            AddParagraph(doc, "第二章、天時節令", fontSize: 16, isBold: true, spacingAfter: 200);
            if (resultsql != null)
            {
                AddParagraph(doc, $"原性：{resultsql}");
            }
            AddParagraph(doc, "");

            // 用神分析
            // 1. 初始化引擎並執行分析
            var engine = new YongShenEngine();
            var result = engine.Analyze(data);

            // 2. 加入章節標題 (參考您之前的樣式)
            AddParagraph(doc, "【依訣八字取用】", fontSize: 16, isBold: true, color: "C00000");

            // 3. 寫入格局與原理
            AddParagraph(doc, $"● 命局格局：{result.PatternName}", fontSize: 12, isBold: true);
            AddParagraph(doc, $"● 取用原理：{result.Logic}", fontSize: 11, spacingAfter: 150);

            // 4. 強調靈魂用神 (使用顯眼的藍色)
            AddParagraph(doc, $"● 靈魂用神：【{result.YongShen}】", fontSize: 14, isBold: true, color: "0070C0");

            // 5. 寫入運勢建議
            AddParagraph(doc, "● 命途吉凶指導：", fontSize: 11, isBold: true);
            AddParagraph(doc, $"  ▶ 吉運趨勢：{result.ProsperityAdvice}", fontSize: 11);

            // 如果有忌諱 (Taboo)，也一併印出
            if (!string.IsNullOrEmpty(result.ProsperityAdvice))
            {
                // 這裡可以根據需要增加對忌諱的描述
            }

            string diag  = DiagnoseDayStemAuspiciousness(data.Bazi.DayPillar.HeavenlyStem, data);

            // 產出到 Word
            AddParagraph(doc, "【日干分析：十干喜忌根性】", fontSize: 16, isBold: true, color: "C00000");
            AddParagraph(doc, diag, fontSize: 11);
            AddAdvancedDiagnosis(doc, data);
            //一柱
            // --- 關鍵整合：調用 YiZhuEngine ---
            var yiZhuEngine = new YiZhuEngine();
            var diagnosis = yiZhuEngine.Diagnose(data,request.Gender); // gender 1男2女

            if (diagnosis != null)
            {
                // 輸出您要求的格式
                AddParagraph(doc, "【日元性情】", fontSize: 14, isBold: true, color: "8B0000");
                AddParagraph(doc, diagnosis.DayMasterAnalysis, fontSize: 11);
                //
                AddParagraph(doc, "元神根基：", fontSize: 12, isBold: true);
                if (!string.IsNullOrEmpty(diagnosis.MarriageStatus)) AddParagraph(doc, diagnosis.MarriageStatus, fontSize: 11);
                if (!string.IsNullOrEmpty(diagnosis.CareerStatus)) AddParagraph(doc, diagnosis.CareerStatus, fontSize: 11);
                if (!string.IsNullOrEmpty(diagnosis.ChildrenStatus)) AddParagraph(doc, diagnosis.ChildrenStatus, fontSize: 11);
                if (!string.IsNullOrEmpty(diagnosis.RelativesAnalysis)) AddParagraph(doc, diagnosis.RelativesAnalysis, fontSize: 11);
                AddParagraph(doc, "");
            }
            // 四柱分析
            AddPillarAnalysis(doc, data);
            //跳頁
            //doc.CreateParagraph().CreateRun().AddBreak(BreakType.PAGE);
            ////跳頁
            doc.CreateParagraph().CreateRun().AddBreak(BreakType.PAGE);
            // --- 三、核心資訊 ---
            AddParagraph(doc, "第三章、先天星盤", fontSize: 16, isBold: true, color: "2E75B6");
            AddParagraph(doc, $"【五行局】：{data.WuXingJuText}");
            AddParagraph(doc, $"【命主星】：{data.MingZhu}");
            AddParagraph(doc, $"【身主星】：{data.ShenZhu}");

            AddParagraph(doc, "【十二宮位深度解析】", fontSize: 14, isBold: true, spacingAfter: 300);

            // 依照紫微斗數標準順序排列
            var palaceOrder = new string[] { "命", "兄", "夫", "子", "財", "疾", "遷", "友", "官", "田", "福", "父" };

            foreach (var targetName in palaceOrder)
            {
                // 修正：使用 Contains 進行模糊匹配，避免 "兄弟" vs "兄弟宮" 抓不到的問題
                var palace = data.palaces.FirstOrDefault(p => p.PalaceName.Contains(targetName));
                if (palace == null) continue;

                // --- 修正星曜抓取欄位 ---
                // 您的 JSON 欄位是 MajorStars, SecondaryStars, SmallStars, YearlyGeneralStars
                // 務必確保全部併入，否則「虎、指、符」等小星不會出現
                var allStars = (palace.MajorStars ?? new List<string>())
                    .Concat(palace.SecondaryStars ?? new List<string>())
                    .Concat(palace.SmallStars ?? new List<string>())
                    .Concat(palace.GoodStars ?? new List<string>()) // 補上流年神煞
                    .Concat(palace.BadStars ?? new List<string>()) // 補上流年神煞
                    .Distinct()
                    .ToList();

                // 如果有長生十二神，也補進去
                if (!string.IsNullOrEmpty(palace.LifeCycleStage)) allStars.Add(palace.LifeCycleStage);
                // 1. 先將亮度字串切割成陣列 (例如 "1+,1+" 變成 ["1+", "1+"])
                // 1. 先將亮度字串切割成陣列 (例如 "1+,1+" 變成 ["1+", "1+"])
                var brightnessList = palace.MainStarBrightness?.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                     .Select(b => b.Trim()).ToList();

                // 2. 組合主星與對應亮度
                string mainStarsStr = (palace.MajorStars?.Any() == true)
                    ? string.Join("、", palace.MajorStars.Select((s, i) =>
                    {
                        // 根據索引取對應的亮度，若無則顯示空字串
                        string b = (brightnessList != null && i < brightnessList.Count) ? brightnessList[i] : "";
                        return $"{s}({b})";
                    })) + $" [{palace.DecadeAgeRange}]"
                    : "無主星" + $" [{palace.DecadeAgeRange}]";
                // 1. 宮位抬頭
                //string mainStarsStr = (palace.MajorStars != null && palace.MajorStars.Any())
                //                        ? string.Join("、", palace.MajorStars)
                //                        : "無主星";

                string title = $"● {palace.PalaceName} (坐{palace.EarthlyBranch})";
                if (mainStarsStr != "無主星") title += $" - 主星：{mainStarsStr}";
                AddParagraph(doc, title, isBold: true, fontSize: 12, spacingAfter: 100);

                // 2. 核心邏輯：從 Dictionary (StarMap) 抓取 113 顆星說明
                // 合併所有星曜：主、輔、小星
                //var allStars = (palace.MajorStars ?? new List<string>())
                //    .Concat(palace.SecondaryStars ?? new List<string>())
                //    .Concat(palace.SmallStars ?? new List<string>())
                //    .Distinct();


                foreach (var starName in allStars)
                {
                    if (StarMap.Dict.TryGetValue(starName, out var def))
                    {
                        var pStar = doc.CreateParagraph();
                        var rStar = pStar.CreateRun();
                        rStar.SetFontFamily("標楷體", FontCharRange.None);
                        rStar.IsBold = true;
                        rStar.SetColor("C00000");

                        // 【關鍵改動】：使用 def.StarName (全名) 而不是 starName (JSON 給的簡稱)
                        // 這樣 JSON 給 "馬"，Word 裡會漂亮地印出 "【天馬】"
                        rStar.SetText($"【{def.StarName}】");

                        var rAttr = pStar.CreateRun();
                        rAttr.FontSize = 10;
                        rAttr.SetColor("666666");

                        // 加上防呆處理，避免有些小星沒有 Element 或 Transform 欄位
                        string attrText = $" ({def.Level})";
                        if (!string.IsNullOrEmpty(def.Element)) attrText += $" - 五行{def.Element}{def.Polar}";
                        if (!string.IsNullOrEmpty(def.Transform)) attrText += $"，化氣為{def.Transform}";
                        rAttr.SetText(attrText);

                        // 輸出長篇「星情本義」
                        var pMean = doc.CreateParagraph();
                        pMean.IndentationLeft = 420;
                        var rMean = pMean.CreateRun();
                        rMean.SetFontFamily("標楷體", FontCharRange.None);
                        rMean.FontSize = 10;
                        rMean.SetText(def.Meaning);
                    }
                }

                // --- 下一階段預留：PostgreSQL styledesc 格局說明 ---
                // 目前先註解掉，不影響編譯，也不需要改 AnalysisService
                /*
                var styleInfo = await _analysisDbService.GetStyleDescAsync(palace.PalaceName, mainStarsStr);
                if (styleInfo != null) {
                    AddParagraph(doc, $"※ 格局鑑定：{styleInfo.Description}", color: "0070C0");
                }
                */

                AddParagraph(doc, "------------------------------------------------------------", color: "D9D9D9", spacingAfter: 200);
            }
            //跳頁
            doc.CreateParagraph().CreateRun().AddBreak(BreakType.PAGE);

            // 1.3 核心格局論斷
            AddParagraph(doc, "審時命局", fontSize: 16, isBold: true, spacingAfter: 200);
            string birthMarkAnalysis = AnalyzeTimeDetails(data.Bazi.TimePillar.EarthlyBranch, request.Gender, request.Minute,request.Hour);
            AddParagraph(doc, birthMarkAnalysis);

            // 【新增】執行核心格局分析
            var patternAnalysisText = await AnalyzeLifePalacePattern(data);
            AddParagraph(doc, patternAnalysisText, color: "8B0000", isItalic: string.IsNullOrEmpty(patternAnalysisText));
            //
            //跳頁
            doc.CreateParagraph().CreateRun().AddBreak(BreakType.PAGE);
            await AddZiWeiDeepAnalysis(doc, data);
        }

        private void AddBackground(XWPFDocument doc, string imagePath)
        {
            if (!File.Exists(imagePath)) return;

            // 1. 透過頁首(Header)設定背景，能確保每一頁都自動出現底紙
            var header = doc.CreateHeader(HeaderFooterType.DEFAULT);

            // 取得或建立頁首段落
            var paragraph = (header.Paragraphs.Count > 0) ? header.Paragraphs[0] : header.CreateParagraph();
            var run = paragraph.CreateRun();

            // 2. 修正 FileStream 與引數錯誤
            // 引數順序應為：(路徑, FileMode, FileAccess)
            using (var fs = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
            {
                // 修正模稜兩可的參考：明確指定 NPOI.XWPF.UserModel.PictureType
                // A4 滿版尺寸設定：寬 595pt (約 5443225 EMU), 高 842pt (約 7701280 EMU)
                run.AddPicture(
                    fs,
                    (int)NPOI.XWPF.UserModel.PictureType.JPEG,
                    "background.jpg",
                    Units.ToEMU(595),
                    Units.ToEMU(842)
                );
            }

            // 3. (進階) 如果要讓圖片「浮在文字下方」且「不佔用頁首空間」
            // NPOI 預設 AddPicture 是 Inline (嵌入式)，會把頁首撐大。
            // 若要完美底紙效果，建議將圖片透明度調高，或直接在 Word 範本中設定背景。
        }
        private void AddSignatureAndStamp(XWPFDocument doc)
        {
            var p = doc.CreateParagraph();
            p.Alignment = ParagraphAlignment.RIGHT;
            var run = p.CreateRun();

            //run.AddBreak();
            //run.SetText("道脈承傳 道德親批");
            //run.FontFamily = "標楷體";
            //run.SetFontFamily("標楷體", FontCharRange.EastAsia);
            //run.FontSize = 16;
            //run.AddBreak();

            // 取得圖片路徑
            string imgDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot", "images");

            // 簽名 (縮小尺寸以免跑版)
            AddImageIfExists(run, Path.Combine(imgDir, "signature.png"), 80, 30);
            // 印章 (緊跟在後)
            AddImageIfExists(run, Path.Combine(imgDir, "玉洞子印.png"), 50, 50);
        }

        // 輔助工具：檢查檔案存在才插入圖片
        private void AddImageIfExists(XWPFRun run, string path, int w, int h)
        {
            if (File.Exists(path))
            {
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    run.AddPicture(fs, (int)NPOI.XWPF.UserModel.PictureType.PNG, Path.GetFileName(path), Units.ToEMU(w), Units.ToEMU(h));
                }
            }
        }
        public class ShiShenIntuitionEngine
        {
            public string Diagnose(AstrologyChartResult chartData)
            {
                var bazi = chartData.Bazi;
                StringBuilder sb = new StringBuilder();

                sb.AppendLine("【十神四柱直觀：心性與外環境診斷】");

                // --- 1. 年柱：祖業與根基 ---
                sb.AppendLine($"● 年柱診斷 ({bazi.YearPillar.HeavenlyStem}{bazi.YearPillar.EarthlyBranch.Substring(0, 1)})：");
                sb.Append(AnalyzePillar("年", bazi.YearPillar.HeavenlyStemLiuShen));

                // --- 2. 月柱：性格、事業與手足 ---
                sb.AppendLine($"\n● 月柱診斷 ({bazi.MonthPillar.HeavenlyStem}{bazi.MonthPillar.EarthlyBranch.Substring(0, 1)})：");
                sb.Append(AnalyzePillar("月", bazi.MonthPillar.HeavenlyStemLiuShen));

                // --- 3. 日支：內心世界與配偶 ---
                // 取得日支的第一個藏干十神
                string dayBranchShiShen = bazi.DayPillar.HiddenStemLiuShen.FirstOrDefault() ?? "";
                sb.AppendLine($"\n● 日支診斷 ({bazi.DayPillar.EarthlyBranch.Substring(0, 1)})：");
                sb.Append(AnalyzePillar("日", dayBranchShiShen));

                // --- 4. 時柱：晚年與子女 ---
                sb.AppendLine($"\n● 時柱診斷 ({bazi.TimePillar.HeavenlyStem}{bazi.TimePillar.EarthlyBranch.Substring(0, 1)})：");
                sb.Append(AnalyzePillar("時", bazi.TimePillar.HeavenlyStemLiuShen));

                return sb.ToString();
            }

            public List<string> AutoAnalyze(AstrologyChartResult chartData)
            {
                var bazi = chartData.Bazi;
                var resultList = new List<string>();

                // 掃描年柱天干
                AddMatchedRule(resultList, "年", "Stem", bazi.YearPillar.HeavenlyStemLiuShen);
                // 掃描月柱天干
                AddMatchedRule(resultList, "月", "Stem", bazi.MonthPillar.HeavenlyStemLiuShen);
                // 掃描日柱地支
                AddMatchedRule(resultList, "日", "Branch", bazi.DayPillar.HeavenlyStemLiuShen);
                // 掃描時柱天干
                AddMatchedRule(resultList, "時", "Stem", bazi.TimePillar.HeavenlyStemLiuShen);

                return resultList;
            }
            public class ShiShenRule
            {
                public string Pillar { get; set; }  // 年/月/日/時
                public string Type { get; set; }    // 天干(Stem)/地支(Branch)
                public string Role { get; set; }    // 正官/傷官...
                public string Description { get; set; }
            }

            public List<string> RunFullDiagnosis(AstrologyChartResult chartData)
            {
                var bazi = chartData.Bazi;
                var results = new List<string>();

                // 1. 診斷四柱天干 (十神直觀)
                AddRule(results, "年", "Stem", bazi.YearPillar.HeavenlyStemLiuShen);
                AddRule(results, "月", "Stem", bazi.MonthPillar.HeavenlyStemLiuShen);
                AddRule(results, "時", "Stem", bazi.TimePillar.HeavenlyStemLiuShen);

                // 2. 診斷日支 (家庭與內心)
                string dayBranch = bazi.DayPillar.EarthlyBranch.Substring(0, 1);
                string dayShiShen = bazi.DayPillar.HeavenlyStemLiuShen;
                AddRule(results, "日", "Branch", dayShiShen);

                // 3. 診斷地支身體 (高級斷要規則)
                var branches = new[] {
                bazi.YearPillar.EarthlyBranch.Substring(0, 1),
                bazi.MonthPillar.EarthlyBranch.Substring(0, 1),
                bazi.DayPillar.EarthlyBranch.Substring(0, 1),
                bazi.TimePillar.EarthlyBranch.Substring(0, 1)
                };

                foreach (var b in branches.Distinct())
                {
                    var healthRule = FullRules.FirstOrDefault(r => r.Type == "Health" && r.Role == b);
                    if (healthRule != null) results.Add($"【身體部位提示】{b}：{healthRule.Description}");
                }

                return results;
            }

            private void AddRule(List<string> results, string pillar, string type, string role)
            {
                var match = FullRules.FirstOrDefault(r => r.Pillar == pillar && r.Type == type && r.Role == role);
                if (match != null) results.Add($"【{pillar}柱直觀】{match.Description}");
            }
            // 專家整理的部分規則列表 (補全至 80 條的基準)

            public static List<ShiShenRule> FullRules = new List<ShiShenRule>
            {
                // --- 正官系列 (1-10) ---
                new ShiShenRule { Pillar = "年", Type = "Stem", Role = "官", Description = "祖上有德，少年得志，易得長輩提攜。" },
                new ShiShenRule { Pillar = "月", Type = "Stem", Role = "官", Description = "為人循規蹈矩，膽小怕事，不願出頭，適合領薪水或在體制內發展。" },
                new ShiShenRule { Pillar = "時", Type = "Stem", Role = "官", Description = "子女端正，晚年有名望，晚景優渥。" },
                new ShiShenRule { Pillar = "日", Type = "Branch", Role = "官", Description = "配偶自律正直，家庭管理嚴謹。" },
                new ShiShenRule { Pillar = "通用", Type = "Character", Role = "官", Description = "重視名譽，自尊心強，行事講求規範。" },

                // --- 七殺系列 (11-20) ---
                new ShiShenRule { Pillar = "年", Type = "Stem", Role = "殺", Description = "早年環境壓力大，性格早熟且倔強，祖業競爭烈。" },
                new ShiShenRule { Pillar = "月", Type = "Stem", Role = "殺", Description = "極度愛面子，好大喜功，說話易誇張、喜歡吹牛。" },
                new ShiShenRule { Pillar = "時", Type = "Stem", Role = "殺", Description = "晚年具威嚴，性格剛毅，但與子女溝通易生隔閡。" },
                new ShiShenRule { Pillar = "日", Type = "Branch", Role = "殺", Description = "配偶性格剛烈，內心危機意識強，不輕易信任他人。" },
                new ShiShenRule { Pillar = "通用", Type = "Character", Role = "殺", Description = "具備爆發力與環境適應力，好動不好靜。" },

                // --- 正印系列 (21-30) ---
                new ShiShenRule { Pillar = "年", Type = "Stem", Role = "印", Description = "出身家風淳樸，早年多得長輩庇蔭照顧。" },
                new ShiShenRule { Pillar = "月", Type = "Stem", Role = "印", Description = "心地仁慈、才華內斂，適合幕僚或教育藝術工作。" },
                new ShiShenRule { Pillar = "時", Type = "Stem", Role = "印", Description = "保守奉獻，晚年生活平穩，不喜往外衝，注重精神寄託。" },
                new ShiShenRule { Pillar = "日", Type = "Branch", Role = "印", Description = "配偶溫和體貼，對日主有奉獻精神。" },
                new ShiShenRule { Pillar = "通用", Type = "Character", Role = "印", Description = "保守、熱愛奉獻，與傷官的往外衝特質相反。" },

                // --- 偏印(梟)系列 (31-40) ---
                new ShiShenRule { Pillar = "年", Type = "Stem", Role = "梟", Description = "祖業外遷，早年生活較孤癖，領悟力特強。" },
                new ShiShenRule { Pillar = "月", Type = "Stem", Role = "梟", Description = "性格孤僻刻薄，易鑽牛角尖，但具備特殊才藝。" },
                new ShiShenRule { Pillar = "時", Type = "Stem", Role = "梟", Description = "晚年易感孤獨，對宗教或玄學有特殊領悟。" },
                new ShiShenRule { Pillar = "日", Type = "Branch", Role = "梟", Description = "配偶精明但沈默，內心世界難以捉摸。" },

                // --- 傷官系列 (41-50) ---
                new ShiShenRule { Pillar = "年", Type = "Stem", Role = "傷", Description = "不論喜忌祖上都不太好，祖業漂流，六親緣薄。" },
                new ShiShenRule { Pillar = "月", Type = "Stem", Role = "傷", Description = "才華橫溢但目無尊長，說話直接易傷人，手足緣薄。" },
                new ShiShenRule { Pillar = "日", Type = "Branch", Role = "傷", Description = "配偶清高自負，婚姻易生波折，總覺得伴侶不如自己。" },
                new ShiShenRule { Pillar = "時", Type = "Stem", Role = "傷", Description = "子緣較薄，晚年具強烈表達慾望，容易才華不遇。" },
                new ShiShenRule { Pillar = "通用", Type = "Character", Role = "傷", Description = "好勝心強，有反叛心理，浪漫主義且幻想力豐富。" },

                // --- 食神系列 (51-60) ---
                new ShiShenRule { Pillar = "年", Type = "Stem", Role = "食", Description = "早年衣食無憂，平安和諧，性格民主。" },
                new ShiShenRule { Pillar = "月", Type = "Stem", Role = "食", Description = "發達之象，能言善道，行事不強加於人，溫和細膩。" },
                new ShiShenRule { Pillar = "時", Type = "Stem", Role = "食", Description = "晚年有福氣，子女有才華且孝順。" },
                new ShiShenRule { Pillar = "日", Type = "Branch", Role = "食", Description = "配偶性格開朗，重視物質與精神享受。" },

                // --- 偏財系列 (61-70) ---
                new ShiShenRule { Pillar = "年", Type = "Stem", Role = "才", Description = "祖上闊綽，早年財運佳，慷慨大方。" },
                new ShiShenRule { Pillar = "月", Type = "Stem", Role = "才", Description = "具備商業開創力，人際廣闊，感情世界豐富。" },
                new ShiShenRule { Pillar = "時", Type = "Stem", Role = "才", Description = "晚年財富穩健，易有意外之財遺留子女。" },

                // --- 比肩系列 (10條) ---
                new ShiShenRule { Pillar = "年", Type = "Stem", Role = "比", Description = "上有兄姊，或出身平民家庭，早年多與同齡人競爭。" },
                new ShiShenRule { Pillar = "月", Type = "Stem", Role = "比", Description = "自我意識強，意志堅定，具團隊合作精神但有時固執己見。" },
                new ShiShenRule { Pillar = "時", Type = "Stem", Role = "比", Description = "晚年具備獨立性，子女有競爭心，自身意志力到老不衰。" },
                new ShiShenRule { Pillar = "日", Type = "Branch", Role = "比", Description = "配偶如友，志趣相投，但也易因各持己見而爭執。" },
                new ShiShenRule { Pillar = "通用", Type = "Character", Role = "比", Description = "意志力堅定，處事果斷，不輕易改變想法。" },

                // --- 劫財系列 (10條) ---
                new ShiShenRule { Pillar = "年", Type = "Stem", Role = "劫", Description = "早年家境易變，易有祖業被他人分享或爭奪之象。" },
                new ShiShenRule { Pillar = "月", Type = "Stem", Role = "劫", Description = "膽大好惹事，好喝嫖賭，社會適應力強但易破財。" },
                new ShiShenRule { Pillar = "時", Type = "Stem", Role = "劫", Description = "晚年開銷大，需防因子女或社交而財散，晚景勞碌。" },
                new ShiShenRule { Pillar = "日", Type = "Branch", Role = "劫", Description = "配偶外向且豪爽，但婚姻中易有財務紛爭或競爭者。" },
                new ShiShenRule { Pillar = "組合", Type = "Logic", Role = "劫傷", Description = "身旺透劫財與傷官：膽大目無法紀，好侵犯人，容易有牢獄之災。" },

                // --- 正財系列 (10條) ---
                new ShiShenRule { Pillar = "年", Type = "Stem", Role = "財", Description = "祖上有經營之才，早年生活規律，家境穩定。" },
                new ShiShenRule { Pillar = "月", Type = "Stem", Role = "財", Description = "行事規矩，重視信用，為人勤儉，適合從事財務或穩定管理。" },
                new ShiShenRule { Pillar = "時", Type = "Stem", Role = "財", Description = "晚年財富穩健，能留產業於子孫，生活安穩。" },
                new ShiShenRule { Pillar = "日", Type = "Branch", Role = "財", Description = "配偶賢慧持家，對日主有實質財富助力。" },

                // --- 補充：傷官/食神 深層解析 (根據附件) ---
                new ShiShenRule { Pillar = "月", Type = "Stem", Role = "食", Description = "發達之象。行事民主不強加於人，溫和細膩，能言善道。" },
                new ShiShenRule { Pillar = "日", Type = "Branch", Role = "傷", Description = "女命日坐傷官：婚姻不順，心高氣傲，自負清高，看誰都不如自己。" },
                new ShiShenRule { Pillar = "組合", Type = "Logic", Role = "印傷", Description = "正印與傷官相反：正印保守奉獻，傷官不保守且主動往外衝。" },

                // --- 補充：偏印(梟) 深層解析 ---
                new ShiShenRule { Pillar = "月", Type = "Stem", Role = "梟", Description = "性格孤僻，處理問題較刻薄，但領悟力極強。" },
                new ShiShenRule { Pillar = "時", Type = "Stem", Role = "梟", Description = "晚年喜安靜，易鑽牛角尖，與子女較疏遠。" },

                // --- 地支健康深度補充 (12條全) ---
                new ShiShenRule { Pillar = "地支", Type = "Health", Role = "丑", Description = "注意脾胃問題、消化系統弱點。" },
                new ShiShenRule { Pillar = "地支", Type = "Health", Role = "申", Description = "注意骨骼健康、骨質疏鬆或關節勞損。" },
                new ShiShenRule { Pillar = "地支", Type = "Health", Role = "辰", Description = "注意闌尾炎、小腸功能。" },

                // --- 地支身體診斷 (71-80，取自《高級斷要》最後一頁) ---
                new ShiShenRule { Pillar = "地支", Type = "Health", Role = "寅", Description = "注意腰部、脊椎、偏頭疼。" },
                new ShiShenRule { Pillar = "地支", Type = "Health", Role = "卯", Description = "注意手指、頸椎、肝臟問題。" },
                new ShiShenRule { Pillar = "地支", Type = "Health", Role = "巳", Description = "注意心臟病、糖尿病預警。" },
                new ShiShenRule { Pillar = "地支", Type = "Health", Role = "午", Description = "注意血液系統、嚴重時防血疾。" },
                new ShiShenRule { Pillar = "地支", Type = "Health", Role = "未", Description = "注意胃部、神經痛。" },
                new ShiShenRule { Pillar = "地支", Type = "Health", Role = "酉", Description = "注意大腸、肺部、痔瘡。" },
                new ShiShenRule { Pillar = "地支", Type = "Health", Role = "戌", Description = "注意血稠、血液循環。" },
                new ShiShenRule { Pillar = "地支", Type = "Health", Role = "亥", Description = "注意前列腺、泌尿系統。" },
                new ShiShenRule { Pillar = "地支", Type = "Health", Role = "子", Description = "注意腎臟、子宮問題。" },
                new ShiShenRule { Pillar = "地支", Type = "Health", Role = "辰", Description = "注意小腸、闌尾炎。" }
            };

            private void AddMatchedRule(List<string> results, string pillar, string type, string shishen)
            {
                var rule = FullRules.FirstOrDefault(r => r.Pillar == pillar && r.Type == type && r.Role == shishen);
                if (rule != null) results.Add($"【{pillar}{type}直觀】{rule.Description}");
            }

            public string GetAutoDiagnosisResult(AstrologyChartResult chartData)
            {
                var bazi = chartData.Bazi;
                var matchedTexts = new List<string>();

                // 1. 匹配四柱天干 (十神表面之象)
                AddMatch(matchedTexts, "年", "Stem", bazi.YearPillar.HeavenlyStemLiuShen);
                AddMatch(matchedTexts, "月", "Stem", bazi.MonthPillar.HeavenlyStemLiuShen);
                AddMatch(matchedTexts, "時", "Stem", bazi.TimePillar.HeavenlyStemLiuShen);

                // 2. 匹配日支 (心性之象)
                string dayBranchShiShen = bazi.DayPillar.HeavenlyStemLiuShen;
                AddMatch(matchedTexts, "日", "Branch", dayBranchShiShen);

                // 3. 匹配地支身體部位 (高級斷要)
                var allBranches = new[] {
                    bazi.YearPillar.EarthlyBranch.Substring(0,1),
                    bazi.MonthPillar.EarthlyBranch.Substring(0,1),
                    bazi.DayPillar.EarthlyBranch.Substring(0,1),
                    bazi.TimePillar.EarthlyBranch.Substring(0,1)
                }.Distinct();

                foreach (var b in allBranches)
                {
                    var health = FullRules.FirstOrDefault(r => r.Type == "Health" && r.Role == b);
                    if (health != null) matchedTexts.Add($"【健康預警】{b}位：{health.Description}");
                }

                // 4. 特殊組合判定 (劫財+傷官)
                string allStems = bazi.YearPillar.HeavenlyStemLiuShen + bazi.MonthPillar.HeavenlyStemLiuShen + bazi.TimePillar.HeavenlyStemLiuShen;
                if (allStems.Contains("劫") && allStems.Contains("傷"))
                {
                    matchedTexts.Add("【行為警示】身旺透劫財與傷官：好惹事生非，膽大目無法紀，需防官司口舌。");
                }

                // 將所有命中的段落串接
                return string.Join("\n", matchedTexts);
            }


            public class StrengthResult
            {
                /// <summary>
                /// 日主總體旺衰得分 (正數為強，負數為弱)
                /// 參考滴天髓：> 80 為旺極，< -80 為弱極
                /// </summary>
                public double Score { get; set; }

                /// <summary>
                /// 日主旺衰狀態描述 (如：身弱、身強、旺極、弱極)
                /// </summary>
                public string Status { get; set; } = string.Empty;

                /// <summary>
                /// 各五行原始得分 (木、火、土、金、水)
                /// </summary>
                public Dictionary<string, double> ElementScores { get; set; } = new Dictionary<string, double>();

                /// <summary>
                /// 月令對日主的影響 (得令、失令)
                /// </summary>
                public string SeasonEffect { get; set; } = string.Empty;

                /// <summary>
                /// 地支是否有根 (滴天髓強調：地支是力量的真實載體)
                /// </summary>
                public bool HasRoot { get; set; }

                /// <summary>
                /// 命局之「病」 (最強或最弱的干擾元素)
                /// </summary>
                public string PrimaryIllness { get; set; } = string.Empty;
            }

            /// <summary>
            /// 滴天髓用神分析結果
            /// </summary>
            public class YongShenResult
            {
                public string PatternName { get; set; } = string.Empty;   // 格局名稱
                public string YongShen { get; set; } = string.Empty;      // 靈魂用神
                public string XiShen { get; set; } = string.Empty;        // 喜神
                public string JiShen { get; set; } = string.Empty;        // 忌神
                public string Logic { get; set; } = string.Empty;         // 取用原理 (病藥論)
                public string ProsperityAdvice { get; set; } = string.Empty; // 吉運建議
            }

            public enum PatternType { 旺極, 強極, 弱極, 正格 }


            public class YongShenEngine
            {
                public YongShenResult Analyze(AstrologyChartResult chartData)
                {
                    // 現在我們統一從 BaziAnalysisResult (ShiShenResult 型別) 讀取數據
                    var analysis = chartData.BaziAnalysisResult;

                    if (analysis == null) return new YongShenResult { Logic = "缺少分析數據" };

                    // 這裡原本傳入的是 StrengthResult，請改為直接傳入分析物件
                    if (analysis.Score >= 80) return GetSpecialPattern(PatternType.旺極, analysis);
                    if (analysis.Score <= -80) return GetSpecialPattern(PatternType.弱極, analysis);

                    return GetNormalPattern(analysis);
                }

                // 修正點：引數型別從 Bazi 改為 BaziInfo (或者若沒用到 bazi，可以直接移除該參數)
                private bool IsWangJi(StrengthResult strength, BaziInfo bazi)
                {
                    return strength.Score >= 80;
                }

                private bool IsRuoJi(StrengthResult strength, BaziInfo bazi)
                {
                    return strength.Score <= -80;
                }

                //private YongShenResult GetSpecialPattern(PatternType type, StrengthResult strength)
                //{
                //    return type switch
                //    {
                //        PatternType.旺極 => new YongShenResult
                //        {
                //            PatternName = "從旺格 (旺極)",
                //            YongShen = "印星、比劫",
                //            Logic = "此為《滴天髓》所述『旺之極者不可損』。全局勢不可擋，必須順其旺勢。",
                //            ProsperityAdvice = "運行印、比運大吉；行財官運觸怒強神大凶。"
                //        },
                //        PatternType.弱極 => new YongShenResult
                //        {
                //            PatternName = "從弱格 (弱極)",
                //            YongShen = "官殺、財星、食傷",
                //            Logic = "日元孤立無氣，必須『棄命從之』。順勢原則：官旺隨官，財旺隨財。",
                //            ProsperityAdvice = "行財、官運亨通富貴；行印、比幫身運必凶。"
                //        },
                //        _ => GetNormalPattern(strength)
                //    };
                //}

                // 修正前：private YongShenResult GetNormalPattern(StrengthResult strength)
                // 修正後：
                private YongShenResult GetNormalPattern(ShiShenResult analysis)
                {
                    if (analysis.Score < 0) // 現在直接讀取 ShiShenResult 裡的 Score
                    {
                        return new YongShenResult
                        {
                            PatternName = "正格身弱",
                            YongShen = "印星、比劫",
                            Logic = "根據口訣：日元力量不足以擔當財官，故取印星生身、比劫幫身為用。",
                            ProsperityAdvice = "運行印、比大運人生亨通。"
                        };
                    }
                    else
                    {
                        return new YongShenResult
                        {
                            PatternName = "正格身強",
                            YongShen = "官殺、食傷、財星",
                            Logic = "日元得勢，取官殺克制、食傷洩秀或財星耗力以達中和。",
                            ProsperityAdvice = "運行財、官、食傷運事業發達。"
                        };
                    }
                }

                private YongShenResult GetSpecialPattern(PatternType type, ShiShenResult analysis)
                {
                    // ... 內文邏輯保持不變，僅修改參數型別為 ShiShenResult ...
                    return type switch
                    {
                        PatternType.旺極 => new YongShenResult        {
                            PatternName = "從旺格 (旺極)",
                            YongShen = "印星、比劫",
                            Logic = "根據口訣所述『旺之極者不可損』。全局勢不可擋，必須順其旺勢。",
                            ProsperityAdvice = "運行印、比運大吉；行財官運觸怒強神大凶。"
                        },
                        PatternType.弱極 => new YongShenResult         {
                            PatternName = "從弱格 (弱極)",
                            YongShen = "官殺、財星、食傷",
                            Logic = "日元孤立無氣，必須『棄命從之』。順勢原則：官旺隨官，財旺隨財。",
                            ProsperityAdvice = "行財、官運亨通富貴；行印、比幫身運必凶。"
                        },
                        _ => GetNormalPattern(analysis)
                    };
                }
            }

            private void AddMatch(List<string> list, string pillar, string type, string role)
            {
                var rule = FullRules.FirstOrDefault(r => r.Pillar == pillar && r.Type == type && r.Role == role);
                if (rule != null) list.Add(rule.Description);
            }

            private string AnalyzeSocialAndWealth(AstrologyChartResult chartData)
            {
                var stems = chartData.Bazi;
                string others = stems.YearPillar.HeavenlyStemLiuShen +
                                stems.MonthPillar.HeavenlyStemLiuShen +
                                stems.TimePillar.HeavenlyStemLiuShen;

                if (others.Contains("劫") && others.Contains("傷"))
                {
                    return "【重要警告】天干透出劫財與傷官，性格好勝且膽大，行事易衝動惹官非，理財宜守不宜攻。";
                }
                if (others.Contains("比"))
                {
                    return "【性格診斷】天干透比肩，意志力強，人際關係多為平輩互動，自尊心重。";
                }
                return "";
            }

            private string AnalyzePillar(string pillarType, string shishen)
            {
                if (string.IsNullOrEmpty(shishen)) return "  - 無顯著十神特徵。\n";

                return shishen switch
                {
                    "官" => pillarType switch
                    {
                        "年" => "  ▶ 祖上有德，出生於書香門第或官宦之家，具備天生的名譽感。\n",
                        "月" => "  ▶ 循規蹈矩，膽小怕事不願出頭，適合打工或在體制內發展。\n",
                        _ => "  ▶ 為人端正，自律性強，重視社會責任與形象。\n"
                    },
                    "殺" => pillarType switch
                    {
                        "月" => "  ▶ 好勝心極強，注重他人評價，好大喜功，甚至有時喜歡吹牛或誇張。\n",
                        "時" => "  ▶ 性格剛毅，晚年仍具威嚴，但需注意與子女溝通之圓融。\n",
                        _ => "  ▶ 聰明伶俐，但性格有時偏激，具備極強的環境適應力。\n"
                    },
                    "印" => pillarType switch
                    {
                        "時" => "  ▶ 思想保守，熱愛奉獻，晚年生活平穩，不喜歡往外衝，注重精神生活。\n",
                        "月" => "  ▶ 文才出眾，性格仁慈，易得長輩或上司之提攜照顧。\n",
                        _ => "  ▶ 性格溫和，有包容力，但有時過於依賴他人，缺乏衝勁。\n"
                    },
                    "梟" => "  ▶ 領悟力強，但性格孤僻，有時處理問題較為刻薄，具備特殊才藝。\n",
                    "財" => "  ▶ 為人現實，重視效率，具備商業頭腦，行事講求回報。\n",
                    "才" => "  ▶ 慷慨大方，人際關係廣闊，性格多情，具備開創性的財富觀。\n",
                    "食" => "  ▶ 願意民主，行事不強加於人，溫和、細膩，具備服務精神與口才。\n",
                    "傷" => pillarType switch
                    {
                        "年" => "  ▶ 祖業漂流，六親緣薄，內心反叛，好勝心強，喜歡追求田園詩般的生活。\n",
                        "月" => "  ▶ 有傷手足或手足失和之象，才華橫溢但容易目無尊長。\n",
                        "日" => "  ▶ 配偶清高自負，婚姻需防口舌，看誰都不如自己。\n",
                        "時" => "  ▶ 子女緣分較薄，或子女個性極強，晚年具備極強的表達慾望。\n",
                        _ => "  ▶ 好奇心強，幻想力豐富，浪漫主義者，不拘小節。\n"
                    },
                    "劫" => "  ▶ 膽大、目無法紀，好惹事生非，社會適應力強但易有破財之象。\n",
                    "比" => "  ▶ 自我意識強，意志堅定，具備團隊合作精神，但也容易固執己見。\n",
                    _ => "  - 待進階解析。\n"
                };
            }
        }

        public class BaziAdvancedAnalysis
        {
            // 判斷五行生扶關係 (新派：辰丑土能生金，未戌土能脆金)
            private bool IsSupport(string stem, string branch)
            {
                string s = stem.Substring(0, 1);
                string b = branch.Substring(0, 1);

                return s switch
                {
                    "辛" or "庚" => "丑辰申酉".Contains(b), // 講義：辰丑生金，未戌脆金不生
                    "壬" or "癸" => "申酉亥子".Contains(b),
                    "甲" or "乙" => "亥子寅卯".Contains(b),
                    "丙" or "丁" => "寅卯巳午".Contains(b),
                    "戊" or "己" => "巳午辰戌丑未".Contains(b),
                    _ => false
                };
            }

            public string EvaluateStrength(AstrologyChartResult chartData)
            {
                var bazi = chartData.Bazi;
                string dayStem = bazi.DayPillar.HeavenlyStem;
                string monthBranch = bazi.MonthPillar.EarthlyBranch.Substring(0, 1);
                string yearBranch = bazi.YearPillar.EarthlyBranch.Substring(0, 1);
                string dayBranch = bazi.DayPillar.EarthlyBranch.Substring(0, 1);

                // 步驟 1: 判定月令對日干的基本關係 (講義規則 2)
                // 辛生巳月 -> 巳火克辛金 -> 基本弱 50%
                bool isMonthSupport = IsSupport(dayStem, monthBranch);

                // 步驟 2: 判定月令受制情況 (講義規則 2, 5)
                // 蔡先生案例：年支卯木(生巳)、日支亥水(沖巳)
                int monthSuppressedCount = 0;

                // 檢查年支與月令關係 (卯木生巳火，不計受制)
                // 檢查日支與月令關係 (亥水沖巳火，計一次受制)
                if (IsSuppressed(monthBranch, dayBranch)) monthSuppressedCount++;

                // 步驟 3: 綜合判定格局
                // 規則：月令克制日干，50%弱。若月令一次受制，仍有50%力量。
                if (!isMonthSupport)
                {
                    if (monthSuppressedCount == 1)
                    {
                        // 檢查是否有其他生扶 (時干戊土為印星生身)
                        bool hasOtherSupport = IsSupport(dayStem, bazi.TimePillar.HeavenlyStem);

                        if (hasOtherSupport)
                            return "身弱格：月令克制日元，受制一次但仍有力，幸得時干印星生身，取印為用。";
                        else
                            return "從弱格：月令克制且無他支生扶。";
                    }
                }
                return "其他格局判定中...";
            }

            // 判斷受制 (沖、泄、克)
        }

        public static class KongWangHelper
        {
            // 根據日柱尋找空亡地支
            public static List<string> GetKongWang(string dayPillar)
            {
                // 這是新派必備的六十甲子空亡表邏輯
                var kongWangMap = new Dictionary<string, List<string>>();

                // 甲子旬 (子丑寅卯辰巳午未申酉戌亥) -> 戌亥空
                string[] xun1 = { "甲子", "乙丑", "丙寅", "丁卯", "戊辰", "己巳", "庚午", "辛未", "壬申", "癸酉" };
                foreach (var s in xun1) kongWangMap[s] = new List<string> { "戌", "亥" };

                // 甲戌旬 -> 申酉空
                string[] xun2 = { "甲戌", "乙亥", "丙子", "丁丑", "戊寅", "己卯", "庚辰", "辛巳", "壬午", "癸未" };
                foreach (var s in xun2) kongWangMap[s] = new List<string> { "申", "酉" };

                // 甲申旬 -> 午未空
                string[] xun3 = { "甲申", "乙酉", "丙戌", "丁亥", "戊子", "己丑", "庚寅", "辛卯", "壬辰", "癸巳" };
                foreach (var s in xun3) kongWangMap[s] = new List<string> { "午", "未" };

                // 甲午旬 -> 辰巳空
                string[] xun4 = { "甲午", "乙未", "丙申", "丁酉", "戊戌", "己亥", "庚子", "辛丑", "壬寅", "癸卯" };
                foreach (var s in xun4) kongWangMap[s] = new List<string> { "辰", "巳" };

                // 甲辰旬 -> 寅卯空 (蔡先生 辛亥日 屬於這一旬)
                string[] xun5 = { "甲辰", "乙巳", "丙午", "丁未", "戊申", "己酉", "庚戌", "辛亥", "壬子", "癸丑" };
                foreach (var s in xun5) kongWangMap[s] = new List<string> { "寅", "卯" };

                // 甲寅旬 -> 子丑空
                string[] xun6 = { "甲寅", "乙卯", "丙辰", "丁巳", "戊午", "己未", "庚申", "辛酉", "壬戌", "癸亥" };
                foreach (var s in xun6) kongWangMap[s] = new List<string> { "子", "丑" };

                return kongWangMap.ContainsKey(dayPillar) ? kongWangMap[dayPillar] : new List<string>();
            }
        }
        public void AddAdvancedDiagnosis(XWPFDocument doc, AstrologyChartResult chartData)
        {
            var bazi = chartData.Bazi;
            string dayStem = bazi.DayPillar.HeavenlyStem;
            string monthBranch = bazi.MonthPillar.EarthlyBranch.Substring(0, 1);

            AddParagraph(doc, "【理法細論】", fontSize: 16, isBold: true, color: "2F5496");

            // 1. 月令分析
            string monthRelation = "克制"; // 此處應根據五行邏輯動態判斷
            AddParagraph(doc, $"● 月令分析：日干【{dayStem}】生於【{monthBranch}】月，月令對日元為{monthRelation}，決定了命局 50% 的基本強度。", fontSize: 11);

            // 2. 作用力診斷
            AddParagraph(doc, "● 作用力診斷：地支見「亥巳相沖」，月令巳火受日支亥水沖擊，力量受損（受制一次）。", fontSize: 11);

            // 3. 格局結論
            string conclusion = "身弱用印";
            string logic = "因辛金生於巳月本弱，月令雖一次受制但克制力仍在，此時全靠時干【戊土】正印貼身生扶，此格局趨向「身弱用印」。";

            AddParagraph(doc, $"● 體用判定：{conclusion}", fontSize: 12, isBold: true, color: "C00000");
            AddParagraph(doc, $"診斷邏輯：{logic}", fontSize: 11);
        }
        private static bool IsSuppressed(string main, string attacker)
        {
            if (string.IsNullOrEmpty(main) || string.IsNullOrEmpty(attacker)) return false;

            // 只取地支首字
            string m = main.Substring(0, 1);
            string a = attacker.Substring(0, 1);

            // 1. 六沖 (最強烈的受制)
            string[,] collisions = { { "子", "午" }, { "丑", "未" }, { "寅", "申" }, { "卯", "酉" }, { "辰", "戌" }, { "巳", "亥" } };
            for (int i = 0; i < 6; i++)
            {
                if ((m == collisions[i, 0] && a == collisions[i, 1]) || (m == collisions[i, 1] && a == collisions[i, 0]))
                    return true;
            }

            // 2. 燥土脆金 (講義規則：未、戌見申、酉，力度等同於火克金)
            if ((m == "申" || m == "酉") && (a == "未" || a == "戌")) return true;

            // 3. 濕土晦火 (講義規則：辰、丑見巳、午，火減力)
            if ((m == "巳" || m == "午") && (a == "辰" || a == "丑")) return true;

            // 4. 地支相刑 (講義定義：相刑雙方均減力)
            // 寅巳刑、戌未刑、丑戌刑
            if ((m == "寅" && a == "巳") || (m == "巳" && a == "寅")) return true;
            if ((m == "戌" && a == "未") || (m == "未" && a == "戌")) return true;
            if ((m == "丑" && a == "戌") || (m == "戌" && a == "丑")) return true;

            // 5. 合絆 (講義規則：合絆雙方均減力九成，是極強的受制)
            // 子丑合、寅亥合、卯戌合、辰酉合、巳申合、午未合
            string[,] bonds = { { "子", "丑" }, { "寅", "亥" }, { "卯", "戌" }, { "辰", "酉" }, { "巳", "申" }, { "午", "未" } };
            for (int i = 0; i < 6; i++)
            {
                if ((m == bonds[i, 0] && a == bonds[i, 1]) || (m == bonds[i, 1] && a == bonds[i, 0]))
                    return true;
            }

            // 6. 泄 (主動生出的字使原字減力)
            // 如：木生火，木減力；火生土，火減力
            if (m == "寅" && (a == "巳" || a == "午")) return true; // 木泄於火
            if (m == "卯" && (a == "巳" || a == "午")) return true;
            if (m == "巳" && (a == "辰" || a == "戌" || a == "丑" || a == "未")) return true; // 火泄於土
            if (m == "午" && (a == "辰" || a == "戌" || a == "丑" || a == "未")) return true;
            if (m == "申" && (a == "子" || a == "亥")) return true; // 金泄於水
            if (m == "酉" && (a == "子" || a == "亥")) return true;
            if (m == "子" && (a == "寅" || a == "卯")) return true; // 水泄於木
            if (m == "亥" && (a == "寅" || a == "卯")) return true;

            return false;
        }
        private string DiagnoseDayStemAuspiciousness(string dayStem, AstrologyChartResult chartData)
        {
            // 取得年、月、時柱的天干環境
            string env = chartData.Bazi.YearPillar.HeavenlyStem +
                         chartData.Bazi.MonthPillar.HeavenlyStem +
                         chartData.Bazi.TimePillar.HeavenlyStem;

            StringBuilder sb = new StringBuilder();
            string formula = ""; // 原始口訣
            List<string> analysis = new List<string>(); // 診斷結論

            switch (dayStem)
            {
                case "甲":
                    formula = "甲木最喜見庚丁，四季日夜要通根，見辰化為青龍貴，見戌魁罡屬福星。午火為死本消滅，癸水侵蝕反傷身。";
                    if (env.Contains("庚") || env.Contains("丁")) analysis.Add("【天顯時才】天干透庚丁，具備修剪與淬煉，主事業有成。");
                    if (env.Contains("癸")) analysis.Add("【水多腐木】癸水過重，需防依賴心強或身體濕氣重。");
                    break;
                case "乙":
                    formula = "乙木天干喜丙陽，丙火午火姓名揚，癸來滋潤未培根，富貴榮華家世昌。亥子齊來隨水去，一遇未庫必蹉跎。";
                    if (env.Contains("丙")) analysis.Add("【向陽花草】見丙火陽光，主才華橫溢，能揚名於外。");
                    if (env.Contains("癸")) analysis.Add("【雨露滋潤】癸水適量，主根基穩固，家世榮昌。");
                    break;
                case "丙":
                    formula = "丙陽尤畏己癸侵，更忌幹頭顯壬辛。己遮癸擋名利敗，辛重壬並禍己身。單甲只寅富貴至，惟恐重重癸水侵。";
                    if (env.Contains("甲")) analysis.Add("【木火通明】單甲入命，主富貴雙全。");
                    if (env.Contains("癸") || env.Contains("己")) analysis.Add("【雲霧遮日】見己癸透干，名利易受阻礙，宜守成。");
                    break;
                case "丁":
                    formula = "丁火甲母不一般，制戊化癸利路寬，酉財亥官能助我，若不富豪也掌權。乙梟藏透皆為忌，丙劫爭奪主凶頑。";
                    if (env.Contains("甲")) analysis.Add("【有母避寒】甲木為母引火，事業根基極深，利路寬廣。");
                    if (env.Contains("丙")) analysis.Add("【奪光之嫌】見丙火劫財，需防競爭者掠奪成果。");
                    break;
                case "戊":
                    formula = "戊土尤愛四庫臨，刑沖財官反為真，辛金傷官能為禍，奪食丙梟是真神。壬癸就我榮華至，殺印相生必成名。";
                    if (env.Contains("壬") || env.Contains("癸")) analysis.Add("【水潤山川】壬癸財來就我，主一生榮華，財源不絕。");
                    if (env.Contains("辛")) analysis.Add("【土金傷官】辛金透干易招口舌非議，宜謹言慎行。");
                    break;
                case "己":
                    formula = "己土喜愛見財官，壬財甲官宜透幹。不喜乙草侵，丙母貼身酉相親。乙丁疊見重重禍，如不早夭一世貧。";
                    if (env.Contains("壬") || env.Contains("甲")) analysis.Add("【財官雙美】壬甲透干，格局純粹，主名利雙收。");
                    if (env.Contains("乙")) analysis.Add("【野草雜亂】乙木貼身侵蝕，主心神不寧，需注意壓力調節。");
                    break;
                case "庚":
                    formula = "庚金皆喜土包藏，壬河丁星利名揚。戊梟重疊難成事，死絕反喜戊來幫。癸傷本是惹禍物，身到死鄉敗一場。";
                    if (env.Contains("壬") || env.Contains("丁")) analysis.Add("【金白水清】壬丁齊見，主名聲遠播，具備領導風範。");
                    if (env.Contains("戊")) analysis.Add("【土多金埋】戊土過重則才華難顯，宜求新求變。");
                    break;
                case "辛":
                    formula = "辛金本來愛壬傷，壬丙交輝定吉昌。酉地子地都為福，只怕庚來劫一場。戊己塵封不為貴，兒孫遭殃貧賤真。";
                    if (env.Contains("壬") && env.Contains("丙")) analysis.Add("【壬丙交輝】極佳格局！主一生吉昌，名利雙收。");
                    if (env.Contains("庚")) analysis.Add("【劫財爭奪】見庚金劫財，需防財來財去，理財宜謹慎。");
                    break;
                case "壬":
                    formula = "壬河本喜戊堤防，逢庚則變活力揚。辰龍吸水到亥門，甲食丙財顯吉祥。運走乙卯一片傷，己官本是榮身物。";
                    if (env.Contains("戊") || env.Contains("庚")) analysis.Add("【堤防有功】見戊土防洪、庚金發源，主具備大將之風。");
                    if (env.Contains("甲") || env.Contains("丙")) analysis.Add("【食能生財】甲丙齊見，主財源廣進，福分深厚。");
                    break;
                case "癸":
                    formula = "癸水最怕見酉金，寒金結露必定凝。見丙見戊堪為美，戊癸一合富貴真。辛金寶貝能作福，己殺有化不夭貧。";
                    if (env.Contains("丙") || env.Contains("戊")) analysis.Add("【水火既濟】見丙戊主貴，若天干戊癸合，更是富貴之象。");
                    if (env.Contains("辛")) analysis.Add("【金水相生】辛金滋扶，主有智慧與福澤。");
                    break;
            }

            sb.AppendLine($"【十干喜忌口訣】：{formula}");
            if (analysis.Count > 0)
            {
                sb.AppendLine("\n【命格即時診斷】：");
                foreach (var item in analysis) sb.AppendLine($"  ▶ {item}");
            }
            else
            {
                //sb.AppendLine("\n【命格即時診斷】：天干尚未透出顯著喜忌之神，需配合地支進階分析。");
            }

            return sb.ToString();
        }
        private string AnalyzeDayStemAffinity(string dayStem, string dayBranch, AstrologyChartResult chartData)
        {
            // 取得所有出現在天干的十神
            string allStems = chartData.Bazi.YearPillar.HeavenlyStem +
                              chartData.Bazi.MonthPillar.HeavenlyStem +
                              chartData.Bazi.TimePillar.HeavenlyStem;

            return dayStem switch
            {
                "甲" => "【甲木喜忌】最喜見庚金、丁火，四季日夜要通根。見辰位為青龍貴，見戌位魁罡屬福星。忌午火死地與癸水過多傷身。",
                "乙" => "【乙木喜忌】天干喜丙火陽光，見午火能揚名。喜癸水滋潤、未土培根。忌庚金利刃過度限制，忌亥子水多隨波逐流。",
                "丙" => "【丙火喜忌】喜單甲或寅木，主富貴。畏己土遮蔽、癸水擋光，更忌辛壬並見，禍及己身。",
                "丁" => "【丁火喜忌】喜甲木為母資生，喜制戊化癸，利路寬廣。喜酉財亥官助我掌權。忌乙木梟印與丙火奪光。",
                "戊" => "【戊土喜忌】尤愛四庫（辰戌丑未）臨命。喜壬癸水為財，榮華至。喜殺印相生（丙、甲）必成名。忌辛金傷官與丙火奪食。",
                "己" => "【己土喜忌】喜壬財、甲官透干。喜丙火貼身。不喜乙木貼身侵擾，若乙丁疊見則易貧困。",
                "庚" => "【庚金喜忌】喜土包藏，喜壬水、丁火洗煉則名利揚。忌戊土重疊掩埋金光，忌癸水傷官惹禍。",
                "辛" => "【辛金喜忌】本愛壬水傷官，壬丙交輝定吉昌。喜酉、子之地為福。怕庚金劫財，忌戊己土過厚封埋金性。",
                "壬" => "【壬水喜忌】喜戊土堤防，逢庚金生助則活力揚。喜辰、亥為門戶，甲食丙財顯吉祥。忌己土混雜使水質混濁。",
                "癸" => "【癸水喜忌】喜見丙火、戊土，戊癸一合主富貴。喜辛金、己土有化。最怕見酉金，寒金結露使水凝結不流。",
                _ => ""
            };
        }

        private void AddPillarAnalysis(XWPFDocument doc, AstrologyChartResult chartData)
        {
                //AddParagraph(doc, "【根苗花果】", fontSize: 16, isBold: true, color: "2F5496");

                var bazi = chartData.Bazi;
                var pillars = new[] {
                    new { Title = "年柱 (祖蔭/根基)", Pillar = bazi.YearPillar.HeavenlyStem + bazi.YearPillar.EarthlyBranch+"-"+bazi.YearPillar.NaYin },
                    new { Title = "月柱 (父母/事業)", Pillar = bazi.MonthPillar.HeavenlyStem + bazi.MonthPillar.EarthlyBranch+"-"+bazi.MonthPillar.NaYin },
                    new { Title = "日柱 (自身/配偶)", Pillar = bazi.DayPillar.HeavenlyStem + bazi.DayPillar.EarthlyBranch+"-"+bazi.DayPillar.NaYin },
                    new { Title = "時柱 (子嗣/晚運)", Pillar = bazi.TimePillar.HeavenlyStem + bazi.TimePillar.EarthlyBranch+"-"+bazi.TimePillar.NaYin }
                 };

            foreach (var p in pillars)
            {
                if (BaziDataRepo.JiaZiDict.TryGetValue(p.Pillar, out var info))
                {
                    AddParagraph(doc, $"● {p.Title}：{p.Pillar}", fontSize: 13, isBold: true);
                    AddParagraph(doc, $"　性質：{info.Nature}。{info.Preference}。", fontSize: 11);
                    AddParagraph(doc, $"　神煞吉凶：{info.Preference+ info.Omen}。", fontSize: 11, color: "7F7F7F");
                }
            }
        }
        public class JiaZiDetail
        {
            public string Nature { get; set; }     // 性質
            public string Preference { get; set; } // 喜忌
            public string Omen { get; set; }       // 吉凶神煞
        }

        public static class BaziDataRepo
        {
            public static readonly Dictionary<string, JiaZiDetail> JiaZiDict = new Dictionary<string, JiaZiDetail>
    {
        { "甲子", new JiaZiDetail { Nature = "金，爲寶物", Preference = "喜金木旺地", Omen = "進神喜，福星，平頭，懸針，破字" } },
        { "乙丑", new JiaZiDetail { Nature = "金，爲頑礦", Preference = "喜火及南方日時", Omen = "福星，華蓋，正印" } },
        { "丙寅", new JiaZiDetail { Nature = "火，爲爐炭", Preference = "喜冬及木", Omen = "福星，祿刑，平頭，聾啞" } },
        { "丁卯", new JiaZiDetail { Nature = "火，爲爐煙", Preference = "喜巽地及秋冬", Omen = "平頭，截路，懸針" } },
        { "戊辰", new JiaZiDetail { Nature = "木，山林山野處不材之木", Preference = "喜水", Omen = "祿庫，華蓋，水馬庫，棒杖，伏神，平頭" } },
        { "己巳", new JiaZiDetail { Nature = "木，山頭花草", Preference = "喜春及秋", Omen = "祿庫，八吉，闕字，曲腳" } },
        { "庚午", new JiaZiDetail { Nature = "土，路旁幹土", Preference = "喜水及春", Omen = "福星，官貴，截路，棒杖，懸針" } },
        { "辛未", new JiaZiDetail { Nature = "土，含萬寶，待秋成", Preference = "喜秋及火", Omen = "華蓋，懸針，破字" } },
        { "壬申", new JiaZiDetail { Nature = "金，戈戟", Preference = "大喜子午卯酉", Omen = "平頭，大敗，妨害，聾啞，破字，懸針" } },
        { "癸酉", new JiaZiDetail { Nature = "金，金之椎鑿", Preference = "喜木及寅卯", Omen = "伏神，破字，聾啞" } },
        { "甲戌", new JiaZiDetail { Nature = "火，火所宿處", Preference = "喜春及夏", Omen = "正印，華蓋，平頭，懸針，破字，棒杖" } },
        { "乙亥", new JiaZiDetail { Nature = "火，火之熱氣", Preference = "喜土及夏", Omen = "天德，曲腳" } },
        { "丙子", new JiaZiDetail { Nature = "水，江湖", Preference = "喜木及土", Omen = "福星，官貴，平頭，聾啞，交神，飛刃" } },
        { "丁丑", new JiaZiDetail { Nature = "水，水之不流清澈處", Preference = "喜金及夏", Omen = "華蓋，進神，平頭，飛刃，闕字" } },
        { "戊寅", new JiaZiDetail { Nature = "土，堤阜城郭", Preference = "喜木及火", Omen = "伏神，俸杖，聾啞" } },
        { "己卯", new JiaZiDetail { Nature = "土，破堤敗城", Preference = "喜申酉及火", Omen = "進神，短夭，九丑，闕字，曲腳，懸針" } },
        { "庚辰", new JiaZiDetail { Nature = "金，錫蠟", Preference = "喜秋及微木", Omen = "華蓋，大敗，棒杖，平頭" } },
        { "辛巳", new JiaZiDetail { Nature = "金，金之幽者，雜沙石", Preference = "喜火及秋", Omen = "天德，福星，官貴，截路，大敗，懸針，曲腳" } },
        { "壬午", new JiaZiDetail { Nature = "木，楊柳幹節", Preference = "喜春夏", Omen = "官貴，九丑，飛刃，平頭，聾啞，懸針" } },
        { "癸未", new JiaZiDetail { Nature = "木，楊柳根", Preference = "喜冬及水，亦宜春", Omen = "正印，華蓋，短夭，伏神，飛刃，破字" } },
        { "甲申", new JiaZiDetail { Nature = "水，甘井", Preference = "喜春及夏", Omen = "破祿馬，截路，平頭，破字，懸針" } },
        { "乙酉", new JiaZiDetail { Nature = "水，陰壑水", Preference = "喜東方及南", Omen = "破祿，短夭，九丑，曲腳，破字，聾啞" } },
        { "丙戌", new JiaZiDetail { Nature = "土，堆阜", Preference = "喜春夏及水", Omen = "天德，華蓋，平頭，聾啞" } },
        { "丁亥", new JiaZiDetail { Nature = "土，平原", Preference = "喜火及木", Omen = "天乙，福星，官貴，德合，平頭" } },
        { "戊子", new JiaZiDetail { Nature = "火，雷也", Preference = "喜水及春夏", Omen = "伏神，短夭，九丑，杖刑，飛刃" } },
        { "己丑", new JiaZiDetail { Nature = "火，電也", Preference = "喜水及春夏", Omen = "華蓋，大敗，飛刃，曲腳，闕字" } },
        { "庚寅", new JiaZiDetail { Nature = "木，松柏幹節", Preference = "喜秋冬", Omen = "破祿馬，相刑，杖刑，聾啞" } },
        { "辛卯", new JiaZiDetail { Nature = "木，松柏之根", Preference = "喜水土及宜春", Omen = "破祿，交神，九丑，懸針" } },
        { "壬辰", new JiaZiDetail { Nature = "水，龍水", Preference = "喜雷電及春夏", Omen = "正印，天德，水祿馬庫，退神，平頭，聾啞" } },
        { "癸巳", new JiaZiDetail { Nature = "水，水之不息，流入海", Preference = "喜亥子", Omen = "天乙，官貴，德合，伏馬，破字，曲腳" } },
        { "甲午", new JiaZiDetail { Nature = "金，百煉精金", Preference = "喜水木土", Omen = "進神，德合，平頭，破字，懸針" } },
        { "乙未", new JiaZiDetail { Nature = "金，爐炭余金", Preference = "喜大火及土", Omen = "華蓋，截路，曲腳，破字" } },
        { "丙申", new JiaZiDetail { Nature = "火，白茅野燒", Preference = "喜秋冬及木", Omen = "平頭，聾啞，大敗，破字，懸針" } },
        { "丁酉", new JiaZiDetail { Nature = "火，鬼神之靈響", Preference = "喜辰戌丑未", Omen = "天乙，喜神，平頭，破字，聾啞，大敗" } },
        { "戊戌", new JiaZiDetail { Nature = "木，蒿艾之枯者", Preference = "喜火及春夏", Omen = "華蓋，大敗，八專，杖刑，截路" } },
        { "己亥", new JiaZiDetail { Nature = "木，蒿艾之茅", Preference = "喜水及春夏", Omen = "闕字，曲腳" } },
        { "庚子", new JiaZiDetail { Nature = "土，土中空者，屋宇也", Preference = "喜木及金", Omen = "木德合，杖刑" } },
        { "辛丑", new JiaZiDetail { Nature = "土，墳墓", Preference = "喜木及火與春", Omen = "華蓋，懸針，闕字" } },
        { "壬寅", new JiaZiDetail { Nature = "金，金之華飾者", Preference = "喜木及微火", Omen = "截路，平頭，聾啞" } },
        { "癸卯", new JiaZiDetail { Nature = "金，環鈕鈐鐸", Preference = "喜盛火及秋", Omen = "貴人，破字，懸針" } },
        { "甲辰", new JiaZiDetail { Nature = "火，燈也", Preference = "喜夜及水，惡晝", Omen = "華蓋，大敗，平頭，破字，懸針" } },
        { "乙巳", new JiaZiDetail { Nature = "火，燈光也", Preference = "尤喜申酉及秋", Omen = "正祿馬，大敗，曲腳，闕字" } },
        { "丙午", new JiaZiDetail { Nature = "火，月輪", Preference = "喜夜及秋，水旺也", Omen = "喜神，羊刃，交神，平頭，聾啞，懸針" } },
        { "丁未", new JiaZiDetail { Nature = "水，火光也", Preference = "喜夜及秋", Omen = "華蓋，羊刃，退神，八專，平頭，破字" } },
        { "戊申", new JiaZiDetail { Nature = "土，秋間田地", Preference = "喜申酉及火", Omen = "福星，伏馬，杖刑，破字，懸針" } },
        { "己酉", new JiaZiDetail { Nature = "土，秋間禾稼", Preference = "喜申酉及冬", Omen = "退神，截路，九丑，闕字，曲腳，破字，聾啞" } },
        { "庚戌", new JiaZiDetail { Nature = "金，刃劍之餘", Preference = "喜微火及木", Omen = "華蓋，杖刑" } },
        { "辛亥", new JiaZiDetail { Nature = "金，鍾鼎實物", Preference = "喜木火及土", Omen = "正祿馬，懸針" } },
        { "壬子", new JiaZiDetail { Nature = "木，傷水多之木", Preference = "喜火土及夏", Omen = "羊刃，九丑，平頭，聾啞" } },
        { "癸丑", new JiaZiDetail { Nature = "木，傷水少之木", Preference = "喜金水及秋", Omen = "華蓋，福星，八專，破字，闕字，羊刃" } },
        { "甲寅", new JiaZiDetail { Nature = "水，雨也", Preference = "喜夏及火", Omen = "正祿馬，福神，八專，平頭，破字，懸針，聾啞" } },
        { "乙卯", new JiaZiDetail { Nature = "水，露也", Preference = "喜水及火", Omen = "建祿，喜神，八專，九刃，曲腳，懸針" } },
        { "丙辰", new JiaZiDetail { Nature = "土，堤岸", Preference = "喜金及木", Omen = "祿庫，正印，華蓋，截路，平頭，聾啞" } },
        { "丁巳", new JiaZiDetail { Nature = "土，土之沮", Preference = "喜火及西北", Omen = "祿庫，平頭，闕字，曲腳" } },
        { "戊午", new JiaZiDetail { Nature = "火，日輪", Preference = "夏則人畏，冬則人愛", Omen = "伏神，羊刃，九丑，棒杖，懸針" } },
        { "己未", new JiaZiDetail { Nature = "火，日光", Preference = "忌夜，亦畏四者", Omen = "福星，華蓋，羊刃，闕字，曲腳，破字" } },
        { "庚申", new JiaZiDetail { Nature = "木，榴花", Preference = "喜夏，不宜秋冬", Omen = "建祿馬，八專，杖刑，破字，懸針" } },
        { "辛酉", new JiaZiDetail { Nature = "木，榴子", Preference = "喜秋及夏", Omen = "建祿，交神，九丑，八專，懸針，聾啞" } },
        { "壬戌", new JiaZiDetail { Nature = "水，海也", Preference = "喜春夏及木", Omen = "華蓋，退神，平頭，聾啞，杖刑" } },
        { "癸亥", new JiaZiDetail { Nature = "水，百川", Preference = "喜金土火", Omen = "伏馬，大敗，破字，截路" } }
    };
        }

        // 【新增】分析命宮主星以取得格局論斷的邏輯
        private async Task<string> AnalyzeLifePalacePattern(AstrologyChartResult data)
        {
            // 1. 找到命宮
            var mingPalace = data.palaces.FirstOrDefault(p => p.PalaceName.Contains("命宮"));
            if (mingPalace == null || !mingPalace.MajorStars.Any())
            {
                return "命宮無主星，為特殊格局，需詳論遷移宮。";
            }

            // 2. 提取命宮主星和宮位索引
            string mainStar = mingPalace.MajorStars.First(); // 以第一顆主星為代表
            int palaceIndex = mingPalace.Index;

            // 3. 呼叫 IAnalysisService 進行資料庫查詢
            try
            {
                var results = await _analysisDbService.GetAllStarStylesAsync(palaceIndex, mainStar);
                var pattern = results.FirstOrDefault();
                if (pattern != null) // && !string.IsNullOrEmpty(pattern.StarDesc)
                {
                    // 4. 如果查詢到結果，回傳描述
                    return "星局 : " + pattern.Gd + pattern.Bd + " 星:" + pattern.StarByYear;

                }
            }
            catch (System.Exception ex)
            {
                // 如果資料庫查詢出錯，回傳錯誤訊息，方便除錯
                return $"查詢格局資料庫時發生錯誤：{ex.Message}";
            }

            // 5. 如果沒有查詢到結果，回傳預設文字
            return "您的命盤格局均衡，未有明顯特殊格局，待後續章節詳解。";
        }

        // 在 AnalysisReportService.cs 中新增一個處理紫微星曜深度解析的方法
        private async Task AddZiWeiDeepAnalysis(XWPFDocument doc, AstrologyChartResult chartData)
        {
            AddParagraph(doc, "第四章：宮星化象", fontSize: 18, isBold: true);
            AddParagraph(doc, "____________________________________________________________", spacingAfter: 300);

            foreach (var palace in chartData.palaces)
            {
                // 1. 取得乾淨的地支 (例如 "未")
                string cleanBranch = palace.EarthlyBranch.Trim().Substring(0, 1);

                // 2. 判斷查詢的星曜字串
                string dbStarSearch;
                if (palace.MajorStars != null && palace.MajorStars.Any())
                {
                    // 有主星時組合主星 (如 "紫破"、"機")
                    dbStarSearch = string.Join("", palace.MajorStars);
                }
                else
                {
                    // 無主星時，根據您的資料庫範例，傳入一個空格 " "
                    dbStarSearch = " ";
                }

                // 3. 呼叫資料庫
                var desc = await _analysisDbService.GetStarStyleDescAsync(dbStarSearch, cleanBranch);

                // 4. 容錯處理：如果組合查不到，嘗試反向 (針對雙星)
                if (desc == null && palace.MajorStars?.Count == 2)
                {
                    string reversedStar = palace.MajorStars[1] + palace.MajorStars[0];
                    desc = await _analysisDbService.GetStarStyleDescAsync(reversedStar, cleanBranch);
                }

                // 5. 輸出到文件
                if (desc != null)
                {
                    string displayName = (dbStarSearch == " ") ? "命無正曜格" : dbStarSearch;
                    AddParagraph(doc, $"● {palace.PalaceName}宮 (坐{cleanBranch}) - 【{displayName}】", fontSize: 14, isBold: true);

                    // 顯示 stardesc (如：命無正曜府相朝垣因人成事)
                    AddParagraph(doc, $"{desc.StarDesc}", fontSize: 11);

                    // 顯示 gd (優勢)
                    if (!string.IsNullOrEmpty(desc.Gd))
                        AddParagraph(doc, $"  ▶ 特性診斷：{desc.Gd}", fontSize: 10, color: "006400");

                    // 顯示 starbyyear (建議)
                    if (!string.IsNullOrEmpty(desc.StarByYear))
                        AddParagraph(doc, $"  ※ 專家建議：{desc.StarByYear}", fontSize: 10, isItalic: true);
                }
                else if (palace.MajorStars != null && palace.MajorStars.Any())
                {
                    // 只有在「有主星但查無資料」時才顯示待補，空宮若查無格局則可選擇隱藏
                    AddParagraph(doc, $"● {palace.PalaceName}宮 (坐{cleanBranch})", fontSize: 14, isBold: true);
                    AddParagraph(doc, $"【{dbStarSearch}】此星曜組合於「{cleanBranch}」位之詳細格局待補。", fontSize: 10, color: "808080");
                }

                AddParagraph(doc, "------------------------------------------------------------", spacingAfter: 150);
                //跳頁
                //doc.CreateParagraph().CreateRun().AddBreak(BreakType.PAGE);
            }
        }

        private string AnalyzeTimeDetails(string timeBranch, int gender, int minute, int hour)
        {
            // 1. 分鐘換算 (0-119分)
            int totalMinutes = (hour % 2 == 1) ? minute : minute + 60;

            // 2. 判定 時初、時中(時正)、時末 (每段 40 分鐘)
            string timeSection = "";
            if (totalMinutes < 40) timeSection = "時初";
            else if (totalMinutes < 80) timeSection = "時中";
            else timeSection = "時末";

            // 3. 判定 1-8 刻與陰陽時 (1-4陰, 5-8陽)
            int totalQuarter = (totalMinutes / 15) + 1;
            bool isUpperFour = totalQuarter <= 4;

            // 核心修正：計算相對刻數 (上四或下四的第幾刻) -> 這決定了「人數」與「位置」
            int relativeQuarter = (totalQuarter > 4) ? totalQuarter - 4 : totalQuarter;
            int personCount = relativeQuarter; // 口訣：第幾刻即幾個人

            // 4. 判定胎記 (男見一：陰時有；女見二：陽時有)
            bool hasMark = (gender == 1 && isUpperFour) || (gender == 2 && !isUpperFour);
            string markLocation = GetMarkLocation(relativeQuarter);

            // 5. 盲派口訣 (辰戌丑未版)
            string mingshuDetail = "";
            if ("辰戌丑未".Contains(timeBranch))
            {
                mingshuDetail = "辰戌丑未四時孤，不妨父母少親疏。";
                if (timeSection == "時中") mingshuDetail += "時正者先亡父。";
                else mingshuDetail += "時初時末先亡母。";
                mingshuDetail += "兄弟無依靠，祖業難守，男宜僧道女宜姑。";
            }
            // ... 其他地支邏輯保持不變 ...

            // 6. 父母像誰與強勢
            bool isYangBranch = "子寅辰午申戌".Contains(timeBranch);
            string lookLike = (gender == 1) ? (isYangBranch ? "像母親" : "像父親") : (isYangBranch ? "像父親" : "像母親");
            string personality = ((gender == 1 && isYangBranch) || (gender == 2 && !isYangBranch)) ? "。為人比較強勢。" : "。";

            // 7. 整合輸出 (確保人數一定會出現)
            StringBuilder sb = new StringBuilder();
            sb.AppendLine($"【審時聞切 · 四時定數】");
            sb.AppendLine($"您生於 {timeBranch}時之{timeSection}（第{totalQuarter}刻）。外貌個性{lookLike}{personality}");

            // 不論是否有胎記，都顯示出生時人數
            string markPrefix = hasMark ? $"印記印證：依古法推算，您在「{markLocation}」應有胎記或疤痕。" : "印記印證：依四時定數，您天生外觀應無明顯胎記。";
            sb.AppendLine($"{markPrefix}根據出生的刻分判定，您出生時母親身邊只有 {personCount} 個人。");

            sb.AppendLine($"\n【時柱斷驗 · 古傳口訣】");
            sb.AppendLine(mingshuDetail);

            return sb.ToString();
        }

        private string GetMarkLocation(int relativeQuarter)
        {
            return relativeQuarter switch { 1 => "臉上", 2 => "身上", 3 => "手上", 4 => "腳上", _ => "身上" };
        }

        private XWPFParagraph AddParagraph(XWPFDocument doc, string text, int fontSize = 12, bool isBold = false, bool isItalic = false, string color = "000000", ParagraphAlignment alignment = ParagraphAlignment.LEFT, int spacingAfter = 200)
        {
            var p = doc.CreateParagraph();
            p.Alignment = alignment;
            p.SpacingAfter = spacingAfter;
            var run = p.CreateRun();
            run.SetText(text);
            run.FontSize = fontSize;
            run.IsBold = isBold;
            run.IsItalic = isItalic;
            run.SetColor(color);
            run.FontFamily = "DFKai-SB"; // 標楷體在系統中的標準名稱
            run.SetFontFamily("標楷體", FontCharRange.EastAsia);
            return p;
        }
    }
}