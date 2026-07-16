using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

namespace Ecanapi.Helpers
{
    // Applies the 玉洞子印 seal as a VML watermark to all body pages of a DOCX.
    // Cover page (first page) and final seal page get no watermark.
    public static class DocxWatermarkHelper
    {
        public static byte[] AddWatermark(byte[] docxBytes, byte[] sealImageBytes)
        {
            try
            {
                var ms = new MemoryStream();
                ms.Write(docxBytes, 0, docxBytes.Length);
                ms.Position = 0;

                const string emptyHeaderXml =
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                    "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr></w:p></w:hdr>";

                using (var wordDoc = WordprocessingDocument.Open(ms, true))
                {
                    var mainPart = wordDoc.MainDocumentPart!;
                    var body = mainPart.Document.Body!;

                    // Add default header with seal image watermark
                    var defaultHdrPart = mainPart.AddNewPart<HeaderPart>();

                    string imgRelId = "";
                    if (sealImageBytes.Length > 0)
                    {
                        var imagePart = defaultHdrPart.AddImagePart(ImagePartType.Png);
                        using (var imgMs = new MemoryStream(sealImageBytes))
                            imagePart.FeedData(imgMs);
                        imgRelId = defaultHdrPart.GetIdOfPart(imagePart);
                    }

                    // VML image watermark: gain="19999" blacklevel="22938f" = Word "Washout" effect
                    string watermarkHeaderXml =
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"" +
                        " xmlns:v=\"urn:schemas-microsoft-com:vml\"" +
                        " xmlns:o=\"urn:schemas-microsoft-com:office:office\"" +
                        " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                        "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr>" +
                        "<w:r><w:rPr><w:noProof/></w:rPr><w:pict>" +
                        "<v:shape id=\"watermark1\" type=\"#_x0000_t75\"" +
                        " style=\"position:absolute;margin-left:0;margin-top:0;width:13cm;height:13cm;" +
                        "z-index:-251656192;mso-position-horizontal:center;" +
                        "mso-position-horizontal-relative:margin;" +
                        "mso-position-vertical:center;mso-position-vertical-relative:margin\"" +
                        " o:allowincell=\"f\">" +
                        $"<v:imagedata r:id=\"{imgRelId}\" o:title=\"seal\" gain=\"19999\" blacklevel=\"22938f\"/>" +
                        "</v:shape></w:pict></w:r></w:p></w:hdr>";

                    using (var s = defaultHdrPart.GetStream(FileMode.Create, FileAccess.Write))
                    {
                        var b = Encoding.UTF8.GetBytes(watermarkHeaderXml);
                        s.Write(b, 0, b.Length);
                    }

                    // Add empty first-page header (no watermark on cover page)
                    var firstHdrPart = mainPart.AddNewPart<HeaderPart>();
                    using (var s = firstHdrPart.GetStream(FileMode.Create, FileAccess.Write))
                    {
                        var b = Encoding.UTF8.GetBytes(emptyHeaderXml);
                        s.Write(b, 0, b.Length);
                    }

                    string defId   = mainPart.GetIdOfPart(defaultHdrPart);
                    string emptyId = mainPart.GetIdOfPart(firstHdrPart);

                    var allParas = body.Descendants<Paragraph>().ToList();

                    // Step 1: Convert the page break before the seal page into a section break.
                    // Seal page identified by paragraph containing "算命的真蹄".
                    int sealIdx = allParas.FindIndex(p => p.InnerText.Contains("算命的真蹄"));
                    if (sealIdx > 0)
                    {
                        for (int i = sealIdx - 1; i >= 0; i--)
                        {
                            bool hasPageBreak = allParas[i].Descendants<Break>()
                                .Any(b => b.Type != null && b.Type.Value == BreakValues.Page);
                            if (!hasPageBreak) continue;
                            var para = allParas[i];
                            var pPr  = para.ParagraphProperties ?? para.PrependChild(new ParagraphProperties());
                            pPr.AppendChild(new SectionProperties(new SectionType { Val = SectionMarkValues.NextPage }));
                            foreach (var run in para.Elements<Run>().ToList()) run.Remove();
                            break;
                        }
                    }

                    // Step 2: Insert a section break paragraph before 【第一章 so that the cover
                    // page + title page (Section 1) have no watermark.
                    // Section 2 (chapters) gets the watermark; Section 3 (seal page) gets none.
                    allParas = body.Descendants<Paragraph>().ToList(); // refresh after step 1
                    var ch1Para = allParas.FirstOrDefault(p => p.InnerText.StartsWith("【第一章"));
                    if (ch1Para != null)
                    {
                        var sectBreak = new Paragraph(new ParagraphProperties(
                            new SectionProperties(new SectionType { Val = SectionMarkValues.NextPage })
                        ));
                        ch1Para.InsertBeforeSelf(sectBreak);
                    }

                    // Identify sections by order: inline sectPr[0]=Section1 end, [1]=Section2 end
                    var inlineSectPrs = body.Descendants<SectionProperties>()
                        .Where(sp => sp.Parent is ParagraphProperties)
                        .ToList();
                    var bodySectPr = body.Elements<SectionProperties>().FirstOrDefault();
                    if (bodySectPr == null) { bodySectPr = new SectionProperties(); body.AppendChild(bodySectPr); }

                    void ClearHeaders(SectionProperties sp) {
                        foreach (var hr in sp.Elements<HeaderReference>().ToList()) hr.Remove();
                        sp.GetFirstChild<TitlePage>()?.Remove();
                    }
                    void SetHeader(SectionProperties sp, string headerId) =>
                        sp.InsertAt(new HeaderReference { Type = HeaderFooterValues.Default, Id = headerId }, 0);

                    if (inlineSectPrs.Count >= 2)
                    {
                        // Section 1 (cover + title pages): no header
                        ClearHeaders(inlineSectPrs[0]);
                        // Section 2 (chapters with watermark)
                        ClearHeaders(inlineSectPrs[1]);
                        SetHeader(inlineSectPrs[1], defId);
                        // Section 3 (seal page): no header
                        ClearHeaders(bodySectPr);
                        SetHeader(bodySectPr, emptyId);
                    }
                    else if (inlineSectPrs.Count == 1)
                    {
                        // Only seal section break found — apply watermark to that section
                        ClearHeaders(inlineSectPrs[0]);
                        SetHeader(inlineSectPrs[0], defId);
                        ClearHeaders(bodySectPr);
                        SetHeader(bodySectPr, emptyId);
                    }
                    else
                    {
                        // Fallback: watermark on all pages except cover (TitlePage)
                        ClearHeaders(bodySectPr);
                        bodySectPr.InsertAt(new TitlePage(), 0);
                        SetHeader(bodySectPr, defId);
                        bodySectPr.InsertAt(new HeaderReference { Type = HeaderFooterValues.First, Id = emptyId }, 0);
                    }

                    mainPart.Document.Save();

                    // 加入 compatibilityMode=15，避免 Word 以「相容模式」開啟
                    var settingsPart = mainPart.DocumentSettingsPart;
                    if (settingsPart != null)
                    {
                        var compat = settingsPart.Settings.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Compatibility>()
                                     ?? settingsPart.Settings.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Compatibility());
                        // 用 raw XML 插入 compatibilityMode 設定（Name 欄位為 enum，不含此值）
                        var cs = new DocumentFormat.OpenXml.OpenXmlUnknownElement("w:compatSetting");
                        cs.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("w:name", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "compatibilityMode"));
                        cs.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("w:uri", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "http://schemas.microsoft.com/office/word"));
                        cs.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("w:val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "15"));
                        compat.AppendChild(cs);
                        settingsPart.Settings.Save();
                    }
                }

                return ms.ToArray();
            }
            catch
            {
                // Fallback: return original bytes if watermark fails
                return docxBytes;
            }
        }
    }
}
