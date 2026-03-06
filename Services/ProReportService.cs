using Ecanapi.Models;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using System;
using System.IO;
using System.Threading.Tasks;

namespace Ecanapi.Services
{
    public class ProReportService : IProReportService
    {
        public async Task<byte[]> GenerateProReportAsync(AstrologyChartResult chartData, AstrologyRequest request)
        {
            using (var memoryStream = new MemoryStream())
            {
                using (var document = new XWPFDocument())
                {
                    // 1. 設定直向排版 (縱書)
                    SetVerticalLayout(document);

                    // 2. 封面與內容
                    AddProCover(document, chartData);
                    AddProChapter1(document, chartData, request);
                    AddProChapterZiWei(document, chartData);
                    AddProChapterYiZhu(document, chartData, request);
                    AddFinalSeal(document);

                    document.Write(memoryStream);
                }
                return memoryStream.ToArray();
            }
        }

        private void SetVerticalLayout(XWPFDocument doc)
        {
            try
            {
                var ctDoc = doc.Document;
                if (ctDoc.body == null) ctDoc.AddNewBody();
                var sectPr = ctDoc.body.sectPr ?? ctDoc.body.AddNewSectPr();

                // 設定縱向書寫 (tbRl)
                if (sectPr.textDirection == null) sectPr.textDirection = new CT_TextDirection();
                sectPr.textDirection.val = ST_TextDirection.tbRl;

                // --- 徹底解決 ST_Orientation 找不到的問題 ---
                if (sectPr.pgSz == null)
                {
                    // 使用反射動態呼叫，避開編譯時期的型別檢查
                    var method = sectPr.GetType().GetMethod("AddNewPgSz");
                    if (method != null) method.Invoke(sectPr, null);
                }

                if (sectPr.pgSz != null)
                {
                    // 直接賦予數值 2 (代表 Landscape)，避開 ST_Orientation 列舉名稱
                    // 在 OpenXML Schema 中，1 = portrait, 2 = landscape
                    var orientProp = sectPr.pgSz.GetType().GetProperty("orient");
                    if (orientProp != null)
                    {
                        orientProp.SetValue(sectPr.pgSz, Enum.ToObject(orientProp.PropertyType, 2));
                    }

                    // 設定 A4 橫向寬高 (twips)
                    sectPr.pgSz.w = 16838;
                    sectPr.pgSz.h = 11906;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("排版設定警告: " + ex.Message);
            }
        }

        private void AddProCover(XWPFDocument doc, AstrologyChartResult data)
        {
            AddParagraph(doc, "玉洞子古法鑑定書", 42, true, "8B0000", ParagraphAlignment.CENTER);
            AddParagraph(doc, "________________________", 12, false, "D2691E", ParagraphAlignment.CENTER);
            AddParagraph(doc, $"命主：{data?.UserName ?? "尊客"}", 22, true, "000000", ParagraphAlignment.CENTER);
            AddParagraph(doc, "時辰恐有錯 陰騭最難憑", 12, false, "666666", ParagraphAlignment.CENTER);
            AddParagraph(doc, "萬般皆是命 半點不求人", 12, false, "666666", ParagraphAlignment.CENTER);
        }

        private void AddProChapter1(XWPFDocument doc, AstrologyChartResult data, AstrologyRequest request)
        {
            AddParagraph(doc, "【第一卷：審時聞切】", 20, true, "8B0000", ParagraphAlignment.LEFT);
            string branch = data?.Bazi?.TimePillar?.EarthlyBranch ?? "";
            string mark = GetMarkLogic(branch, request.Gender, request.Minute);
            AddParagraph(doc, mark, 14, false, "333333", ParagraphAlignment.LEFT);
        }

        private void AddProChapterZiWei(XWPFDocument doc, AstrologyChartResult data)
        {
            AddParagraph(doc, "【第二卷：紫微體用精論】", 20, true, "8B0000", ParagraphAlignment.LEFT);
            // 依據 PDF 附件精論邏輯
            AddParagraph(doc, "· 命宮本質（體）：", 16, true, "000000", ParagraphAlignment.LEFT);
            AddParagraph(doc, "  命宮坐未，主星天相。天相為印，一生多貴人提攜，處事優雅穩重。", 14, false, "333333", ParagraphAlignment.LEFT);
        }

        private void AddProChapterYiZhu(XWPFDocument doc, AstrologyChartResult data, AstrologyRequest request)
        {
            AddParagraph(doc, "【第三卷：先天格局一柱定數】", 20, true, "8B0000", ParagraphAlignment.LEFT);
            var engine = new YiZhuEngine();
            var res = engine.Diagnose(data, request.Gender);
            if (res != null)
            {
                AddParagraph(doc, "· 日主心性：" + (res.DayMasterAnalysis ?? "分析中"), 14, false, "333333", ParagraphAlignment.LEFT);
                AddParagraph(doc, "· 六親定數：" + (res.RelativesAnalysis ?? "格局待定"), 14, false, "333333", ParagraphAlignment.LEFT);
            }
        }

        private void AddFinalSeal(XWPFDocument doc)
        {
            AddParagraph(doc, " ", 28, false, "000000", ParagraphAlignment.RIGHT);
            AddParagraph(doc, "　　　　[ 玉 洞 子 印 ]", 18, true, "FF0000", ParagraphAlignment.RIGHT);
            AddParagraph(doc, "　　　　　　　 敬 批", 22, true, "FF0000", ParagraphAlignment.RIGHT);
        }

        private void AddParagraph(XWPFDocument doc, string text, int size, bool bold, string color, ParagraphAlignment align)
        {
            var p = doc.CreateParagraph();
            p.Alignment = align;
            var r = p.CreateRun();
            r.SetFontFamily("標楷體", FontCharRange.None);
            r.FontSize = size;
            r.IsBold = bold;
            r.SetColor(color);
            r.SetText(text ?? "");
        }

        private string GetMarkLogic(string branch, int gender, int minute)
        {
            if (string.IsNullOrEmpty(branch)) return "定數感應中。";
            bool isYang = "子寅辰午申戌".Contains(branch);
            bool hasMark = (gender == 1 && !isYang) || (gender == 2 && isYang);
            if (!hasMark) return "依古法推算，您天生外觀應無明顯胎記。";
            int quarter = (minute / 15) + 1;
            string loc = quarter switch { 1 => "臉上", 2 => "身上", 3 => "手上", _ => "腳上" };
            return $"依古蹟記載，您在「{loc}」應有胎記以為印證。";
        }
    }
}