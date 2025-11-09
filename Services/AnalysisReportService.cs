using Ecanapi.Models;
using Ecanapi.Models.Analysis;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.UserModel.Charts;
using NPOI.XWPF.UserModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

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
            using (var memoryStream = new MemoryStream())
            {
                using (var document = new XWPFDocument())
                {
                    AddCoverPage(document, chartData);
                    await AddChapter1(document, chartData, request); // 改為 await
                    document.Write(memoryStream);
                }

                return memoryStream.ToArray();
            }
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
            doc.CreateParagraph().CreateRun().AddBreak(BreakType.PAGE);
        }

        private async Task AddChapter1(XWPFDocument doc, AstrologyChartResult data, AstrologyRequest request)
        {
            AddParagraph(doc, "第一章：先天格局與本命概論", fontSize: 20, isBold: true, color: "00008B", spacingAfter: 400);

            // 1.1 八字命盤
            AddParagraph(doc, "一、八字命盤", fontSize: 16, isBold: true, spacingAfter: 200);
            double[] skyTotal = new double[11];
            var baziTable = doc.CreateTable(6, 5);
            baziTable.Width = 6000;
            baziTable.SetColumnWidth(0, 1000);
            string[] headers = { "", "時柱", "日柱", "月柱", "年柱" };
            for (int i = 0; i < headers.Length; i++)
            {
                var cell = baziTable.GetRow(0).GetCell(i);
                cell.SetText(headers[i]);
                if (cell.GetCTTc().tcPr == null) cell.GetCTTc().AddNewTcPr();
                cell.GetCTTc().tcPr.AddNewTcW().w = "6000";
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

            // 新增干支五行生剋
            // 1.4 干支五行生剋星剎分析
            AddParagraph(doc, "四、干支五行生剋星剎分析", fontSize: 16, isBold: true, spacingAfter: 200);
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
            var strQuery = "SELECT CONCAT(rgcz,xgfx,aqfx,syfx,cyfx,jkfx) AS \"Value\" FROM public.六十甲子命主" + " where rgz = ";
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
            strQuery = "SELECT whw as \"Value\" FROM public.五行 " + " where wh = ";
            strQuery = strQuery + " '" + strFiveKind + "' ";
            resultsql = await _analysisDbService.ExecuteRawQueryAsync(strQuery);
            AddParagraph(doc, " 五行分析:", fontSize: 12, isBold: true, spacingAfter: 200);
            if (resultsql != null)
            {
                AddParagraph(doc, $" {resultsql}");
            }
            if (strFiveKind == "木")
            {
                resultsql = @"木生春月旺、水多主災殃、不是貧困守、
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
                resultsql = @"火到春見陽、木盛火必強、水旺金不旺、
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
                resultsql = @"春月生戊己、難抵甲和乙、用金必有利、
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
                resultsql = @"金生春月地、焉能用甲乙、用土還傷火、
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
                resultsql = @"春水不當權、木盛奪水源、用金禮相見、
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

            // --- 六十甲子日對時 ---
            strQuery = "SELECT \"desc\" as \"Value\" FROM public.\"六十甲子日對時\" " + " where \"Sky\" =  ";
            strQuery = strQuery + " '" + data.Bazi.DayPillar.HeavenlyStem + "' and \"Month\" =  ";
            strQuery = strQuery + " '" + data.Bazi.MonthPillar.EarthlyBranch + "月' and \"time\" like";
            strQuery = strQuery + " '%" + data.Bazi.TimePillar.EarthlyBranch + "%' ";
            resultsql = await _analysisDbService.ExecuteRawQueryAsync(strQuery);
            //Task<IEnumerable<SixtyJiaziDayToHour>> GetSixtyJiaziDayToHourAsync(string? sky, string? month, string? time);
            //var resultMain = await _analysisDbService.GetSixtyJiaziDayToHourAsync(data.Bazi.DayPillar.HeavenlyStem, data.Bazi.MonthPillar.EarthlyBranch, data.Bazi.TimePillar.EarthlyBranch);
            //var pt = resultMain.FirstOrDefault();
            AddParagraph(doc, "二、出生時節", fontSize: 16, isBold: true, spacingAfter: 200);
            if (resultsql != null)
            {
                AddParagraph(doc, $"原性：{resultsql}");
            }
            AddParagraph(doc, "");

            // 1.2 核心資訊
            AddParagraph(doc, "三、核心資訊", fontSize: 16, isBold: true, spacingAfter: 200);
            AddParagraph(doc, $"五行局：{data.WuXingJuText}");
            AddParagraph(doc, $"命主：{data.MingZhu}");
            AddParagraph(doc, $"身主：{data.ShenZhu}");
            AddParagraph(doc, "");

            // 1.3 核心格局論斷
            AddParagraph(doc, "四、核心格局論斷", fontSize: 16, isBold: true, spacingAfter: 200);
            string birthMarkAnalysis = AnalyzeBirthMark(data.Bazi.TimePillar.EarthlyBranch, request.Gender, request.Minute);
            AddParagraph(doc, birthMarkAnalysis);

            // 【新增】執行核心格局分析
            var patternAnalysisText = await AnalyzeLifePalacePattern(data);
            AddParagraph(doc, patternAnalysisText, color: "8B0000", isItalic: string.IsNullOrEmpty(patternAnalysisText));
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
                    return "格局 : " + pattern.Gd + pattern.Bd + " 星:" + pattern.StarByYear;

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


        private string AnalyzeBirthMark(string timeBranch, int gender, int minute)
        {
            bool isYangHour = "子寅辰午申戌".Contains(timeBranch);
            bool isYinHour = "丑卯巳未酉亥".Contains(timeBranch);
            bool hasMark = (gender == 1 && isYinHour) || (gender == 2 && isYangHour);
            if (!hasMark) return "依四時定數，您天生外觀應無明顯胎記。";
            string genderText = gender == 1 ? "男孩" : "女孩";
            string timeTypeText = isYinHour ? "陰時" : "陽時";
            string markLocation = "";
            int quarter = (minute / 15) + 1;
            if (minute >= 60) quarter = 4;
            switch (quarter)
            {
                case 1: markLocation = "臉上"; break;
                case 2: markLocation = "身上"; break;
                case 3: markLocation = "手上"; break;
                case 4: markLocation = "腳上"; break;
            }
            return $"審時聞切，四時定數。您為{genderText}，生於{timeBranch}時({timeTypeText})，此為天定印記之時。依古法推算，您在「{markLocation}」應有胎記或疤痕以為印證。";
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
            return p;
        }
    }
}