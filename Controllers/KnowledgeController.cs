using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Ecanapi.Data;
using Ecanapi.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Globalization;
using System.Security.Claims;
using System.Text;
using System.Text.RegularExpressions;

namespace Ecanapi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Authorize]
    public class KnowledgeController : ControllerBase
    {
        private readonly ApplicationDbContext _db;
        private readonly IConfiguration _config;

        public KnowledgeController(ApplicationDbContext db, IConfiguration config)
        {
            _db = db;
            _config = config;
        }

        private bool IsAdmin()
        {
            var email = User.FindFirstValue(ClaimTypes.Email);
            return email == _config["Admin:Email"];
        }

        // GET /api/Knowledge/rules?page=1&pageSize=20&search=xxx&category=xxx&subcategory=xxx
        [HttpGet("rules")]
        public async Task<IActionResult> GetRules(
            [FromQuery] int page = 1,
            [FromQuery] int pageSize = 20,
            [FromQuery] string? search = null,
            [FromQuery] string? category = null,
            [FromQuery] string? subcategory = null,
            [FromQuery] bool? activeOnly = true)
        {
            if (!IsAdmin()) return Forbid();

            var query = _db.FortuneRules.AsQueryable();
            if (activeOnly == true) query = query.Where(r => r.IsActive);
            if (!string.IsNullOrWhiteSpace(category)) query = query.Where(r => r.Category == category);
            if (!string.IsNullOrWhiteSpace(subcategory)) query = query.Where(r => r.Subcategory == subcategory);
            if (!string.IsNullOrWhiteSpace(search))
            {
                query = query.Where(r =>
                    (r.Title != null && r.Title.Contains(search)) ||
                    r.ResultText.Contains(search) ||
                    (r.ConditionText != null && r.ConditionText.Contains(search)) ||
                    (r.Tags != null && r.Tags.Contains(search)));
            }

            var total = await query.CountAsync();
            var items = await query
                .OrderBy(r => r.Category)
                .ThenBy(r => r.SortOrder)
                .ThenBy(r => r.Id)
                .Skip((page - 1) * pageSize)
                .Take(pageSize)
                .ToListAsync();

            return Ok(new { total, page, pageSize, items });
        }

        // POST /api/Knowledge/rules
        [HttpPost("rules")]
        public async Task<IActionResult> CreateRule([FromBody] FortuneRuleDto dto)
        {
            if (!IsAdmin()) return Forbid();

            var rule = new FortuneRule
            {
                Category = dto.Category,
                Subcategory = dto.Subcategory,
                Title = dto.Title,
                ConditionText = dto.ConditionText,
                ResultText = dto.ResultText,
                SourceFile = dto.SourceFile,
                Tags = dto.Tags,
                IsActive = dto.IsActive ?? true,
                SortOrder = dto.SortOrder ?? 0,
                CreatedAt = DateTime.UtcNow,
                UpdatedAt = DateTime.UtcNow
            };
            _db.FortuneRules.Add(rule);
            await _db.SaveChangesAsync();
            return Ok(rule);
        }

        // PUT /api/Knowledge/rules/{id}
        [HttpPut("rules/{id}")]
        public async Task<IActionResult> UpdateRule(int id, [FromBody] FortuneRuleDto dto)
        {
            if (!IsAdmin()) return Forbid();

            var rule = await _db.FortuneRules.FindAsync(id);
            if (rule == null) return NotFound();

            rule.Category = dto.Category;
            rule.Subcategory = dto.Subcategory;
            rule.Title = dto.Title;
            rule.ConditionText = dto.ConditionText;
            rule.ResultText = dto.ResultText;
            rule.SourceFile = dto.SourceFile;
            rule.Tags = dto.Tags;
            rule.IsActive = dto.IsActive ?? rule.IsActive;
            rule.SortOrder = dto.SortOrder ?? rule.SortOrder;
            rule.UpdatedAt = DateTime.UtcNow;
            await _db.SaveChangesAsync();
            return Ok(rule);
        }

        // DELETE /api/Knowledge/rules/{id}
        [HttpDelete("rules/{id}")]
        public async Task<IActionResult> DeleteRule(int id)
        {
            if (!IsAdmin()) return Forbid();

            var rule = await _db.FortuneRules.FindAsync(id);
            if (rule == null) return NotFound();
            _db.FortuneRules.Remove(rule);
            await _db.SaveChangesAsync();
            return Ok(new { message = "deleted" });
        }

        // GET /api/Knowledge/documents
        [HttpGet("documents")]
        public async Task<IActionResult> GetDocuments()
        {
            if (!IsAdmin()) return Forbid();
            var docs = await _db.KnowledgeDocuments
                .OrderByDescending(d => d.UploadedAt)
                .ToListAsync();
            return Ok(docs);
        }

        // POST /api/Knowledge/upload - parse file and return preview (no DB write)
        [HttpPost("upload")]
        [RequestSizeLimit(50 * 1024 * 1024)]
        [ApiExplorerSettings(IgnoreApi = true)]
        public async Task<IActionResult> Upload([FromForm] IFormFile file, [FromForm] string category, [FromForm] string? subcategory, [FromForm] string? parseMode)
        {
            if (!IsAdmin()) return Forbid();
            if (file == null || file.Length == 0) return BadRequest(new { message = "no file" });

            var ext = Path.GetExtension(file.FileName).ToLowerInvariant();
            var allowed = new[] { ".csv", ".txt", ".docx", ".doc", ".xlsx", ".xls" };
            if (!allowed.Contains(ext))
                return BadRequest(new { message = $"unsupported format: {ext}" });

            using var ms = new MemoryStream();
            await file.CopyToAsync(ms);
            ms.Position = 0;

            List<ParsedRule> parsed;
            try
            {
                parsed = ext switch
                {
                    ".csv" => ParseCsv(ms, file.FileName, category, subcategory),
                    ".txt" => parseMode == "chapters"
                        ? ParseTxtChapters(ms, file.FileName, category)
                        : ParseTxt(ms, file.FileName, category, subcategory),
                    ".docx" => parseMode == "chapters"
                        ? ParseDocxChapters(ms, file.FileName, category)
                        : ParseDocx(ms, file.FileName, category, subcategory, parseMode),
                    ".doc" => ParseDoc(ms, file.FileName, category, subcategory),
                    ".xlsx" => ParseXlsx(ms, file.FileName, category, subcategory, false),
                    ".xls" => ParseXlsx(ms, file.FileName, category, subcategory, true),
                    _ => new List<ParsedRule>()
                };
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = $"parse error: {ex.Message}" });
            }

            return Ok(new
            {
                fileName = file.FileName,
                fileType = ext.TrimStart('.'),
                category,
                subcategory,
                count = parsed.Count,
                preview = parsed.Take(10).ToList(),
                all = parsed
            });
        }

        // POST /api/Knowledge/import - write parsed rules to DB
        [HttpPost("import")]
        public async Task<IActionResult> Import([FromBody] ImportRequest req)
        {
            if (!IsAdmin()) return Forbid();
            if (req.Rules == null || req.Rules.Count == 0)
                return BadRequest(new { message = "no rules" });

            var email = User.FindFirstValue(ClaimTypes.Email);
            var now = DateTime.UtcNow;
            var rules = req.Rules.Select(r => new FortuneRule
            {
                Category = r.Category,
                Subcategory = r.Subcategory,
                Title = r.Title,
                ConditionText = r.ConditionText,
                ResultText = r.ResultText,
                SourceFile = r.SourceFile,
                Tags = r.Tags,
                IsActive = true,
                SortOrder = 0,
                CreatedAt = now,
                UpdatedAt = now
            }).ToList();

            _db.FortuneRules.AddRange(rules);

            var doc = new KnowledgeDocument
            {
                FileName = req.FileName,
                FileType = req.FileType,
                Category = req.Category,
                ContentPreview = req.Rules.FirstOrDefault()?.ResultText?.Length > 300
                    ? req.Rules.First().ResultText[..300]
                    : req.Rules.FirstOrDefault()?.ResultText,
                RuleCount = rules.Count,
                Status = "imported",
                UploadedAt = now,
                UploadedBy = email
            };
            _db.KnowledgeDocuments.Add(doc);

            await _db.SaveChangesAsync();
            return Ok(new { imported = rules.Count, docId = doc.Id });
        }

        // GET /api/Knowledge/categories
        [HttpGet("categories")]
        public async Task<IActionResult> GetCategories()
        {
            if (!IsAdmin()) return Forbid();
            var cats = await _db.FortuneRules
                .Where(r => r.IsActive)
                .Select(r => new { r.Category, r.Subcategory })
                .Distinct()
                .ToListAsync();
            return Ok(cats);
        }

        // ---- Parser helpers ----

        private static List<ParsedRule> ParseCsv(Stream stream, string fileName, string category, string? subcategory)
        {
            var rules = new List<ParsedRule>();
            try
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                var bytes = ReadAllBytes(stream);
                var enc = DetectEncoding(bytes);
                var text = enc.GetString(bytes);
                var lines = text.Split('\n', StringSplitOptions.RemoveEmptyEntries);
                if (lines.Length == 0) return rules;

                // Detect if the file is multi-column (has real data in col 2+)
                // by checking whether any non-first-line has a non-empty second column
                bool isMultiColumn = lines.Skip(1).Take(10).Any(l => {
                    var c = SplitCsvLine(l);
                    return c.Count > 1 && !string.IsNullOrWhiteSpace(c[1]);
                });

                // Skip header only in multi-column files where first row looks like a header
                int start = 0;
                if (isMultiColumn)
                {
                    var firstCols = SplitCsvLine(lines[0]);
                    bool firstRowIsHeader = firstCols.Count > 1 && !string.IsNullOrWhiteSpace(firstCols[1])
                        && !firstCols[0].EndsWith("。") && !firstCols[0].EndsWith("：");
                    if (firstRowIsHeader) start = 1;
                }

                for (int i = start; i < lines.Length; i++)
                {
                    var cols = SplitCsvLine(lines[i]);
                    if (cols.Count == 0) continue;

                    string col0 = cols[0].Trim().TrimEnd('\r');
                    if (string.IsNullOrWhiteSpace(col0)) continue;

                    string title;
                    string result;

                    if (isMultiColumn)
                    {
                        // Multi-column: col0=title, remaining cols=result
                        var rest = cols.Skip(1).Select(c => c.Trim().TrimEnd('\r')).Where(c => c.Length > 0).ToList();
                        title = col0;
                        result = rest.Count > 0 ? string.Join("；", rest) : col0;
                        if (rest.Count > 0 && result == col0) title = "";
                    }
                    else
                    {
                        // Single-column (with or without trailing comma): entire col0 is the rule
                        title = "";
                        result = col0;
                    }

                    if (string.IsNullOrWhiteSpace(result)) continue;

                    rules.Add(new ParsedRule
                    {
                        Category = category,
                        Subcategory = subcategory,
                        Title = string.IsNullOrWhiteSpace(title) ? null : title,
                        ResultText = result,
                        SourceFile = fileName
                    });
                }
            }
            catch
            {
                // ignore parse errors
            }
            return rules;
        }

        private static List<ParsedRule> ParseTxt(Stream stream, string fileName, string category, string? subcategory)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var bytes = ReadAllBytes(stream);
            var enc = DetectEncoding(bytes);
            var text = enc.GetString(bytes);
            return ParseTextContent(text, fileName, category, subcategory);
        }

        // Specialized parser for chapter-structured TXT files (e.g., 八字直斷.txt).
        // Handles three content types:
        //   1. Chapter/section headers (第X章/節) → set Subcategory
        //   2. Verse groups (consecutive 16-char lines) → merged under preceding header
        //   3. Numbered direct-judgment items (N)、or N、) → each becomes one rule
        //   4. Short title (ends with ：, <=25 chars) followed by content → Title + ResultText
        private static List<ParsedRule> ParseTxtChapters(Stream stream, string fileName, string category)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var bytes = ReadAllBytes(stream);
            var enc = DetectEncoding(bytes);
            var text = enc.GetString(bytes).Replace("\r\n", "\n").Replace("\r", "\n");
            var lines = text.Split('\n')
                .Select(l => l.Trim())
                .ToList();

            var rules = new List<ParsedRule>();
            string currentSubcategory = "";
            string? currentTitle = null;
            var versePending = new List<string>();   // accumulate consecutive 16-char verse lines
            var bodyPending = new List<string>();    // accumulate body lines under a short title

            void FlushVerses()
            {
                if (versePending.Count == 0) return;
                rules.Add(new ParsedRule
                {
                    Category = category,
                    Subcategory = currentSubcategory,
                    Title = currentTitle,
                    ResultText = string.Join("\n", versePending),
                    SourceFile = fileName
                });
                currentTitle = null;
                versePending.Clear();
            }

            void FlushBody()
            {
                if (bodyPending.Count == 0 && currentTitle == null) return;
                var result = string.Join("\n", bodyPending).Trim();
                if (result.Length < 3) { bodyPending.Clear(); currentTitle = null; return; }
                rules.Add(new ParsedRule
                {
                    Category = category,
                    Subcategory = currentSubcategory,
                    Title = currentTitle,
                    ResultText = result,
                    SourceFile = fileName
                });
                currentTitle = null;
                bodyPending.Clear();
            }

            for (int i = 0; i < lines.Count; i++)
            {
                var line = lines[i];
                if (string.IsNullOrWhiteSpace(line)) continue;

                // --- Chapter/section header: 第X章 / 第X節 ---
                if (Regex.IsMatch(line, @"第[一二三四五六七八九十百\d]+[章節]"))
                {
                    FlushVerses();
                    FlushBody();
                    currentSubcategory = line;
                    currentTitle = null;
                    continue;
                }

                // --- Numbered direct-judgment items: "1、xxx" / "1)、xxx" / "(1)xxx" ---
                var numberedMatch = Regex.Match(line, @"^[\(（]?\d+[\)）]?[、\.]");
                if (numberedMatch.Success)
                {
                    FlushVerses();
                    FlushBody();
                    var content = line[numberedMatch.Length..].Trim();
                    // Look ahead: if next line is a continuation (no number, decent length), merge it
                    while (i + 1 < lines.Count)
                    {
                        var next = lines[i + 1].Trim();
                        if (string.IsNullOrWhiteSpace(next)) break;
                        bool nextIsNumbered = Regex.IsMatch(next, @"^[\(（]?\d+[\)）]?[、\.]");
                        bool nextIsChapter = Regex.IsMatch(next, @"第[一二三四五六七八九十百\d]+[章節]");
                        bool nextIsShortTitle = next.EndsWith("：") && next.Length <= 25;
                        if (nextIsNumbered || nextIsChapter || nextIsShortTitle) break;
                        if (next.Length > 0 && !Regex.IsMatch(next, @"^\d+[、\.]")) { content += "\n" + next; i++; }
                        else break;
                    }
                    if (!string.IsNullOrWhiteSpace(content))
                        rules.Add(new ParsedRule
                        {
                            Category = category,
                            Subcategory = currentSubcategory,
                            Title = currentTitle,
                            ResultText = (numberedMatch.Value.TrimEnd('、', '.') + " " + content).Trim(),
                            SourceFile = fileName
                        });
                    continue;
                }

                // --- Verse line: exactly 14-18 chars containing punctuation (歌訣) ---
                bool isVerse = line.Length >= 14 && line.Length <= 20
                    && (line.Contains("，") || line.Contains("。") || line.Contains("、"))
                    && !line.Contains("：");
                if (isVerse)
                {
                    FlushBody();
                    versePending.Add(line);
                    continue;
                }

                // --- Short title ending with ：(<= 25 chars) → start of a named group ---
                if (line.EndsWith("：") && line.Length <= 25)
                {
                    FlushVerses();
                    FlushBody();
                    currentTitle = line.TrimEnd('：').Trim();
                    continue;
                }

                // --- Everything else: content paragraph ---
                FlushVerses();
                if (line.Length >= 8)
                {
                    FlushBody();
                    rules.Add(new ParsedRule
                    {
                        Category = category,
                        Subcategory = currentSubcategory,
                        Title = currentTitle,
                        ResultText = line,
                        SourceFile = fileName
                    });
                    currentTitle = null;
                }
                else if (line.Length >= 3)
                {
                    // Too short to be a standalone rule — treat as a sub-title candidate
                    FlushBody();
                    currentTitle = line;
                }
            }

            FlushVerses();
            FlushBody();
            return rules;
        }

        private static List<ParsedRule> ParseDoc(Stream stream, string fileName, string category, string? subcategory)
        {
            // Extract readable text from .doc binary (OLE2 / Word 97-2003).
            // Scans for UTF-16LE character runs (the main text storage format in .doc).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var bytes = ReadAllBytes(stream);
            var text = ExtractDocText(bytes);
            if (string.IsNullOrWhiteSpace(text))
                text = DetectEncoding(bytes).GetString(bytes); // fallback: raw decode
            return ParseTextContent(text, fileName, category, subcategory);
        }

        private static string ExtractDocText(byte[] bytes)
        {
            // Heuristic: scan for UTF-16LE Chinese/ASCII text runs (>= 3 chars).
            // .doc stores main text as UTF-16LE in the WordDocument stream.
            var sb = new StringBuilder();
            int i = 0;
            while (i < bytes.Length - 1)
            {
                // Try UTF-16LE run
                var runSb = new StringBuilder();
                int j = i;
                while (j < bytes.Length - 1)
                {
                    ushort ch = (ushort)(bytes[j] | (bytes[j + 1] << 8));
                    // Accept CJK Unified, Basic Latin printable, CJK punctuation
                    if ((ch >= 0x4E00 && ch <= 0x9FFF) ||   // CJK Unified
                        (ch >= 0x3000 && ch <= 0x303F) ||   // CJK Symbols
                        (ch >= 0xFF00 && ch <= 0xFFEF) ||   // Fullwidth
                        (ch >= 0x0020 && ch <= 0x007E) ||   // ASCII printable
                        ch == 0x000A || ch == 0x000D)        // newline
                    {
                        runSb.Append((char)ch);
                        j += 2;
                    }
                    else break;
                }
                if (runSb.Length >= 3)
                {
                    sb.Append(runSb);
                    i = j;
                }
                else i++;
            }
            // Clean up: remove excessive spaces/control chars
            var result = Regex.Replace(sb.ToString(), @"[\x00-\x08\x0B\x0C\x0E-\x1F]", "");
            result = Regex.Replace(result, @" {3,}", " ");
            return result;
        }

        // Shared text parser: handles both plain paragraphs and structured "title + body" entries.
        // Detects the "年干" structure (short title line followed by content lines).
        private static List<ParsedRule> ParseTextContent(string text, string fileName, string category, string? subcategory)
        {
            var rules = new List<ParsedRule>();

            // Normalize line endings
            text = text.Replace("\r\n", "\n").Replace("\r", "\n");
            var allLines = text.Split('\n').Select(l => l.Trim()).ToList();
            var nonEmpty = allLines.Where(l => l.Length > 0).ToList();

            // Pattern 1: FLAT LIST — each non-empty line is already a complete rule.
            // Detected when most lines contain content punctuation (：、，) even if short.
            // Example: "天機：科學儀器、機械" — a single complete rule per line.
            int linesWithContent = nonEmpty.Count(l => l.Contains("：") || l.Contains("、") || l.Contains("，"));
            bool isFlatList = nonEmpty.Count >= 5 && (double)linesWithContent / nonEmpty.Count >= 0.5;

            // Pattern 2: STRUCTURED title+body — short pure-label lines (<=12 chars, no content
            // punctuation) followed by longer body paragraphs. Example: 四化干性.txt
            var pureTitleLines = nonEmpty.Where(l => l.Length <= 12
                && !l.Contains("，") && !l.Contains("。")
                && !l.Contains("；") && !l.Contains("：") && !l.Contains("、")).ToList();
            bool isStructured = !isFlatList && pureTitleLines.Count >= 3 && nonEmpty.Count > 4;

            if (isFlatList)
            {
                // Every non-empty line = one rule.
                // Special colon-split patterns for 職業類 files:
                //   Pattern A: line has "：：" → split at "：：", left=title(職業), right=result(星)
                //   Pattern B: line has ONE "：" AND first char is digit → left=title(職業), right=result(星)
                //   Otherwise: whole line = resultText
                string? currentSection = null;
                foreach (var line in nonEmpty)
                {
                    // Section headers: no content punctuation at all
                    bool isSection = !line.Contains("：") && !line.Contains("、")
                        && !line.Contains("，") && line.Length <= 20;
                    if (isSection) { currentSection = line; continue; }

                    string? title = null;
                    string result;

                    if (line.Contains("：："))
                    {
                        // Pattern A: 職業：：星組合 → DB: Title=星, ResultText=職業
                        var idx = line.IndexOf("：：", StringComparison.Ordinal);
                        result = line[..idx].Trim();   // 職業預測
                        title  = line[(idx + 2)..].Trim(); // 星組合
                    }
                    else if (line.Contains("：") && line.Length > 0 && char.IsDigit(line[0]))
                    {
                        // Pattern B: 1、職業名稱：星組合 → DB: Title=星, ResultText=職業
                        var idx = line.IndexOf("：", StringComparison.Ordinal);
                        result = line[..idx].Trim();   // 職業預測（含序號）
                        title  = line[(idx + 1)..].Trim(); // 星組合
                    }
                    else
                    {
                        result = line;
                    }

                    if (string.IsNullOrWhiteSpace(result)) result = line;

                    rules.Add(new ParsedRule
                    {
                        Category = category,
                        Subcategory = subcategory ?? currentSection,
                        Title = string.IsNullOrWhiteSpace(title) ? null : title,
                        ResultText = result,
                        SourceFile = fileName
                    });
                }
            }
            else if (isStructured)
            {
                // Parse as structured title+body pairs
                string? currentTitle = null;
                var currentBody = new StringBuilder();

                void FlushStructured()
                {
                    var body = currentBody.ToString().Trim();
                    if (body.Length < 3) { currentBody.Clear(); return; }
                    rules.Add(new ParsedRule
                    {
                        Category = category,
                        Subcategory = subcategory,
                        Title = currentTitle,
                        ResultText = body,
                        SourceFile = fileName
                    });
                    currentTitle = null;
                    currentBody.Clear();
                }

                foreach (var line in allLines)
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    bool looksLikeTitle = line.Length <= 12
                        && !line.Contains("，") && !line.Contains("。")
                        && !line.Contains("；") && !line.Contains("：")
                        && !line.Contains("、");

                    if (looksLikeTitle && currentBody.Length > 0)
                    {
                        FlushStructured();
                        currentTitle = line;
                    }
                    else if (looksLikeTitle && currentBody.Length == 0)
                    {
                        currentTitle = line;
                    }
                    else
                    {
                        if (currentBody.Length > 0) currentBody.Append('\n');
                        currentBody.Append(line);
                    }
                }
                FlushStructured();
            }
            else
            {
                // Parse as double-newline-separated paragraphs
                var paragraphs = Regex.Split(text, @"\n\n+")
                    .Select(p => p.Trim())
                    .Where(p => p.Length > 10)
                    .ToList();

                if (paragraphs.Count <= 1)
                {
                    // Fallback: single newline split
                    paragraphs = allLines.Where(l => l.Length > 10).ToList();
                }

                foreach (var para in paragraphs)
                {
                    var lines = para.Split('\n');
                    var firstLine = lines[0].Trim();
                    var title = firstLine.Length <= 60 ? firstLine : null;
                    var body = title != null && lines.Length > 1
                        ? string.Join(" ", lines.Skip(1).Select(l => l.Trim()).Where(l => l.Length > 0))
                        : para;
                    if (string.IsNullOrWhiteSpace(body)) body = para;

                    rules.Add(new ParsedRule
                    {
                        Category = category,
                        Subcategory = subcategory,
                        Title = title,
                        ResultText = body.Trim(),
                        SourceFile = fileName
                    });
                }
            }

            return rules;
        }

        // Skip styles that represent chart/table layouts rather than narrative text.
        private static readonly HashSet<string> SkipStyles = new(StringComparer.OrdinalIgnoreCase)
            { "af1", "af2", "af3", "af4", "af5" };

        // Specialized parser for chapter-structured .docx files (e.g., 八字直斷核心預測規則指南.docx).
        // Structure: 第X章 header → Subcategory; non-bullet paragraph → ConditionText (section title);
        // each bullet point (NumberingProperties != null) → one independent rule.
        private static List<ParsedRule> ParseDocxChapters(Stream stream, string fileName, string category)
        {
            var rules = new List<ParsedRule>();
            using var doc = WordprocessingDocument.Open(stream, false);
            var body = doc.MainDocumentPart?.Document?.Body;
            if (body == null) return rules;

            // Remove trailing footnote reference numbers: "...成家 1。" → "...成家"
            static string Clean(string s)
                => Regex.Replace(s.Trim(), @"\s*\d+[。\.]\s*$", "").Trim();

            // Collect paragraphs in document order with bullet flag.
            // Also descend into tables (treat each non-empty cell as a paragraph-like token).
            var tokens = new List<(string Text, bool IsBullet)>();

            bool IsBulletParagraph(Paragraph p)
            {
                // Has list numbering definition → bullet or numbered list item
                if (p.ParagraphProperties?.NumberingProperties != null) return true;
                var styleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "";
                // Some templates name bullet styles "ListBullet", "List Bullet", etc.
                return styleId.Contains("List", StringComparison.OrdinalIgnoreCase)
                    || styleId.Contains("Bullet", StringComparison.OrdinalIgnoreCase);
            }

            foreach (var element in body.ChildElements)
            {
                if (element is Paragraph para)
                {
                    var t = Clean(para.InnerText);
                    if (!string.IsNullOrWhiteSpace(t))
                        tokens.Add((t, IsBulletParagraph(para)));
                }
                else if (element is DocumentFormat.OpenXml.Wordprocessing.Table tbl)
                {
                    foreach (var row in tbl.Elements<TableRow>())
                    {
                        // Treat each non-empty cell as a bullet-style item
                        foreach (var cell in row.Elements<TableCell>())
                        {
                            // Merge all paragraphs in cell into one text block
                            var cellText = string.Join(" ", cell.Elements<Paragraph>()
                                .Select(p => p.InnerText.Trim())
                                .Where(t => !string.IsNullOrWhiteSpace(t)));
                            cellText = Clean(cellText);
                            if (!string.IsNullOrWhiteSpace(cellText))
                                tokens.Add((cellText, true)); // table cells = bullet-equivalent
                        }
                    }
                }
            }

            string currentSubcategory = "";
            string currentCondition = "";   // section title / condition label (e.g., "子、午、卯、酉時")

            foreach (var (text, isBullet) in tokens)
            {
                // Chapter/section header: 第X章 / 第X節 / 第X篇
                if (Regex.IsMatch(text, @"第[一二三四五六七八九十百\d]+[章節篇]"))
                {
                    currentSubcategory = text;
                    currentCondition = "";
                    continue;
                }

                if (isBullet)
                {
                    // Each bullet = one rule; use currentCondition as Title + ConditionText
                    rules.Add(new ParsedRule
                    {
                        Category = category,
                        Subcategory = currentSubcategory,
                        Title = string.IsNullOrWhiteSpace(currentCondition) ? null : currentCondition,
                        ConditionText = string.IsNullOrWhiteSpace(currentCondition) ? null : currentCondition,
                        ResultText = text,
                        SourceFile = fileName
                    });
                }
                else
                {
                    // Non-bullet, non-chapter paragraph = section condition/title for following bullets
                    // (e.g., "子、午、卯、酉時" or "甲乙時")
                    currentCondition = text.TrimEnd('：').Trim();
                }
            }

            return rules;
        }

        private static List<ParsedRule> ParseDocx(Stream stream, string fileName, string category, string? subcategory, string? parseMode = null)
        {
            var rules = new List<ParsedRule>();
            using var doc = WordprocessingDocument.Open(stream, false);
            var body = doc.MainDocumentPart?.Document?.Body;
            if (body == null) return rules;

            // "paragraphs" mode: every paragraph becomes one rule; 3-level heading inheritance.
            // Level 1 (Heading1 / 第X章節) → Subcategory
            // Level 2 (Heading2 / bold short line) → Title
            // Level 3 (Heading3+) → ConditionText
            // Content paragraphs inherit all three running values.
            if (parseMode == "paragraphs")
            {
                // Detect heading level (1/2/3). Text patterns take priority over Word styles
                // so that mis-styled headings in Chinese docs are still classified correctly.
                static int HeadingLevel(Paragraph p, string text)
                {
                    // 1. Chinese chapter/section keywords (highest priority, override any style)
                    if (Regex.IsMatch(text, @"^第[一二三四五六七八九十百\d]+[章篇部]")) return 1;
                    if (Regex.IsMatch(text, @"^第[一二三四五六七八九十百\d]+節")) return 2;

                    var styleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "";

                    // 2. Named heading styles: "Heading1", "Heading 1", "标题1", "標題1"
                    var hm = Regex.Match(styleId, @"(?:[Hh]eading\s*|标题\s*|標題\s*)(\d+)");
                    if (hm.Success) return Math.Min(int.Parse(hm.Groups[1].Value), 3);

                    // 3. Outline level from paragraph properties (0-based, 9 = body text)
                    var ol = p.ParagraphProperties?.OutlineLevel?.Val?.Value;
                    if (ol.HasValue && ol.Value < 9) return Math.Min((int)ol.Value + 1, 3);

                    // 4. Font size from runs (half-points: 32=16pt, 28=14pt, 26=13pt)
                    int maxSz = p.Descendants<FontSize>()
                        .Select(fs => { int.TryParse(fs.Val?.Value, out int sz); return sz; })
                        .DefaultIfEmpty(0).Max();
                    if (maxSz >= 32) return 1;
                    if (maxSz >= 28) return 2;
                    if (maxSz >= 26) return 3;

                    // 5. List indent level from numbering (ilvl 0→level2, 1→level3)
                    var ilvl = p.ParagraphProperties?.NumberingProperties?.NumberingLevelReference?.Val?.Value;
                    if (ilvl.HasValue)
                    {
                        if (ilvl.Value == 0) return 2;
                        if (ilvl.Value >= 1) return 3;
                    }

                    // 6. Bold detection (majority of runs are bold)
                    var runs = p.Elements<Run>().ToList();
                    int boldRuns = runs.Count(r => r.RunProperties?.Bold != null
                        && !(r.RunProperties.Bold.Val?.Value == false));
                    bool isParagraphBold = runs.Count > 0 && boldRuns >= runs.Count / 2;

                    // 7. Chinese numeral section (一、xxx) → level 2 if bold or very short
                    bool isCnNumSection = Regex.IsMatch(text, @"^[一二三四五六七八九十]+[、。：]")
                        && text.Length <= 60;
                    if (isCnNumSection && (isParagraphBold || text.Length <= 20)) return 2;

                    // 8. Bold short line fallback → level 2
                    if (isParagraphBold && text.Length <= 50
                        && !text.Contains("。") && !text.Contains("，")) return 2;

                    return 0;
                }

                string? curSub  = subcategory;
                string? curTitle = null;
                string? curCond  = null;

                foreach (var p in body.Elements<Paragraph>())
                {
                    var text = p.InnerText.Trim();
                    if (string.IsNullOrWhiteSpace(text)) continue;
                    var styleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "";
                    if (SkipStyles.Contains(styleId)) continue;

                    int level = HeadingLevel(p, text);

                    if (level == 1)
                    {
                        // Split "第一章 十干斷易篇" → curSub="第一章", curTitle="十干斷易篇"
                        var chapterMatch = Regex.Match(text, @"^(第[一二三四五六七八九十百\d]+[章節篇部])\s*(.*)$");
                        if (chapterMatch.Success)
                        {
                            curSub   = chapterMatch.Groups[1].Value.Trim();
                            curTitle = chapterMatch.Groups[2].Value.Trim();
                            if (string.IsNullOrWhiteSpace(curTitle)) curTitle = null;
                        }
                        else
                        {
                            curSub   = text;
                            curTitle = null;
                        }
                        curCond = null;
                        continue; // heading only updates state, no rule row
                    }
                    else if (level == 2)
                    {
                        curTitle = text;
                        curCond  = null;
                        continue;
                    }
                    else if (level == 3)
                    {
                        curCond = text;
                        continue;
                    }

                    // Content paragraph: save as rule with inherited heading context.
                    rules.Add(new ParsedRule
                    {
                        Category      = category,
                        Subcategory   = curSub,
                        Title         = curTitle,
                        ConditionText = curCond,
                        ResultText    = text,
                        SourceFile    = fileName
                    });
                }
                return rules;
            }

            var paragraphs = body.Elements<Paragraph>()
                .Select(p => new {
                    Text = p.InnerText.Trim(),
                    StyleId = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "",
                    IsBold = p.Descendants<Bold>().Any()
                })
                .Where(p => !string.IsNullOrWhiteSpace(p.Text) && !SkipStyles.Contains(p.StyleId))
                .ToList();

            int numberedCount = paragraphs.Count(p => Regex.IsMatch(p.Text, @"^\d+[\.、]"));

            // Detect "格局" style: short lines ending with ： act as titles, followed by multi-paragraph content.
            // e.g., "紫府同宮格：" (6 chars) → title, next N paragraphs → body merged as one rule.
            int colonTitles = paragraphs.Count(p => p.Text.EndsWith("：") && p.Text.Length <= 20);
            bool isColonGrouped = numberedCount < 5 && colonTitles >= 3;

            if (isColonGrouped)
            {
                string? currentTitle = null;
                string? chapterSubCol = subcategory;
                var bodyParts = new List<string>();

                void FlushColonGroup()
                {
                    if (currentTitle == null && bodyParts.Count == 0) return;
                    var resultText = string.Join("\n", bodyParts).Trim();
                    if (resultText.Length < 3 && currentTitle == null) { bodyParts.Clear(); return; }
                    if (string.IsNullOrWhiteSpace(resultText)) resultText = currentTitle ?? "";
                    rules.Add(new ParsedRule
                    {
                        Category = category,
                        Subcategory = chapterSubCol,
                        Title = currentTitle,
                        ResultText = resultText,
                        SourceFile = fileName
                    });
                    currentTitle = null;
                    bodyParts.Clear();
                }

                foreach (var p in paragraphs)
                {
                    var text = p.Text;
                    // Chapter/section header → update subcategory tracking
                    if (Regex.IsMatch(text, @"^第[一二三四五六七八九十百千\d]+[章節篇]"))
                    {
                        FlushColonGroup();
                        chapterSubCol = text;
                        continue;
                    }
                    // Short line ending with ： → new title
                    if (text.EndsWith("：") && text.Length <= 20)
                    {
                        FlushColonGroup();
                        currentTitle = text.TrimEnd('：').Trim();
                    }
                    else
                    {
                        bodyParts.Add(text);
                    }
                }
                FlushColonGroup();
                return rules;
            }

            // Detect flat-list docx: most paragraphs are self-contained rules containing "："
            // e.g., "擎羊星入命宮：直接、衝動、刀子嘴" — each line is one rule regardless of length.
            int colonLines = paragraphs.Count(p => p.Text.Contains("："));
            bool isFlatListDocx = numberedCount < 5
                && colonLines > 0
                && (double)colonLines / paragraphs.Count >= 0.55;

            if (isFlatListDocx)
            {
                string? currentSub = subcategory;
                foreach (var p in paragraphs)
                {
                    var text = p.Text;
                    if (!text.Contains("："))
                    {
                        // Chapter/section header → update subcategory (chapter headers always override)
                        bool isChapterHeader = Regex.IsMatch(text, @"^第[一二三四五六七八九十百千\d]+[章節篇]");
                        currentSub = isChapterHeader ? text : (string.IsNullOrWhiteSpace(subcategory) ? text : subcategory);
                        continue;
                    }
                    var colonIdx = text.IndexOf('：');
                    var title = text[..colonIdx].Trim();
                    var result = text[(colonIdx + 1)..].Trim();
                    if (string.IsNullOrWhiteSpace(result)) result = text;
                    rules.Add(new ParsedRule
                    {
                        Category = category,
                        Subcategory = currentSub,
                        Title = string.IsNullOrWhiteSpace(title) ? null : title,
                        ResultText = result,
                        SourceFile = fileName
                    });
                }
                return rules;
            }

            // Detect "title + numbered sub-items" structure (e.g., 生年四化入十二宮.docx)
            // If many paragraphs start with digits (2.xxx, 3.xxx...), group them under the preceding non-numbered line.
            if (numberedCount >= 5)
            {
                string? sectionSub = subcategory;
                string? currentTitle = null;
                var bodyParts = new List<string>();

                void FlushGroup()
                {
                    if (bodyParts.Count == 0 && currentTitle == null) return;
                    var resultText = bodyParts.Count > 0
                        ? string.Join("\n", bodyParts)
                        : currentTitle ?? "";
                    if (resultText.Length < 3) { bodyParts.Clear(); currentTitle = null; return; }
                    rules.Add(new ParsedRule
                    {
                        Category = category,
                        Subcategory = sectionSub,
                        Title = currentTitle,
                        ResultText = resultText,
                        SourceFile = fileName
                    });
                    currentTitle = null;
                    bodyParts.Clear();
                }

                foreach (var p in paragraphs)
                {
                    var text = p.Text;
                    // Chapter header: 第X章/節/篇 → update sectionSub and reset title
                    if (Regex.IsMatch(text, @"^第[一二三四五六七八九十百千\d]+[章節篇]"))
                    {
                        FlushGroup();
                        sectionSub = text;
                        continue;
                    }
                    // Section header: must be "一、〈...〉" (with angle brackets), not "一、「...」"
                    bool isSectionHeader = Regex.IsMatch(text, @"^[一二三四五六七八九十百千]+[、。][〈《]");
                    // Numbered sub-item: e.g., "2.xxx" or "2、xxx"
                    bool isNumberedItem = Regex.IsMatch(text, @"^\d+[\.、]");
                    // Chinese numeral sub-point inside a note (e.g., "一、「少小看夫妻」：...")
                    bool isCnNumSub = !isSectionHeader && Regex.IsMatch(text, @"^[一二三四五六七八九十]+[、。]");

                    if (isSectionHeader)
                    {
                        FlushGroup();
                        // Extract palace name from "一、〈命宮〉" or "一、〈命宮 / 事業宮〉"
                        var m = Regex.Match(text, @"[〈《]([^〉》]+)[〉》]");
                        sectionSub = m.Success ? m.Groups[1].Value : text;
                    }
                    else if (isNumberedItem || isCnNumSub)
                    {
                        // Numbered sub-item or Chinese numeral note sub-point → accumulate
                        bodyParts.Add(text);
                    }
                    else
                    {
                        // Skip separator / divider lines (e.g., "----", "====")
                        if (Regex.IsMatch(text, @"^[-=─━─\-]{3,}$")) continue;

                        // Annotation lines (〈注：...〉) → append to current group instead of starting new rule
                        if (text.StartsWith("〈注") || text.StartsWith("〈說明") || text.StartsWith("※") || text.StartsWith("〈附"))
                        {
                            bodyParts.Add(text);
                            continue;
                        }

                        // New rule start line: e.g., "生年祿入命：1.主「福」。..."
                        FlushGroup();
                        // If the line already contains the first sub-item "1." embedded, split it
                        var firstItemMatch = Regex.Match(text, @"1[\.、]");
                        if (firstItemMatch.Success && firstItemMatch.Index > 0)
                        {
                            currentTitle = text[..firstItemMatch.Index].TrimEnd('：').Trim();
                            var firstItem = text[firstItemMatch.Index..].Trim();
                            if (firstItem.Length > 0) bodyParts.Add(firstItem);
                        }
                        else
                        {
                            currentTitle = text.TrimEnd('：').Trim();
                        }
                    }
                }
                FlushGroup();
                return rules;
            }

            // Standard mode: heading/bold lines as titles, substantial paragraphs as rules
            string? stdSub = subcategory;
            string? stdTitle = null;
            var pendingShortLines = new List<string>();

            foreach (var p in paragraphs)
            {
                var text = p.Text;
                bool isHeading = p.StyleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)
                              || p.StyleId.StartsWith("heading", StringComparison.OrdinalIgnoreCase);
                bool isChapterHeader = Regex.IsMatch(text, @"^第[一二三四五六七八九十百千\d]+[章節篇]");
                bool looksLikeTitle = isHeading
                    || (p.IsBold && text.Length <= 30)
                    || Regex.IsMatch(text, @"^([一二三四五六七八九十百千]+[、。〈《]|第[一二三四五六七八九十\d]+[章節篇])");

                if (isChapterHeader)
                {
                    // Flush pending lines before updating chapter subcategory
                    if (pendingShortLines.Count > 0)
                    {
                        var joined = string.Join("；", pendingShortLines);
                        if (joined.Length >= 5)
                            rules.Add(new ParsedRule { Category = category, Subcategory = stdSub,
                                Title = stdTitle, ResultText = joined, SourceFile = fileName });
                        pendingShortLines.Clear();
                    }
                    stdSub = text;
                    stdTitle = null;
                    continue;
                }

                if (looksLikeTitle)
                {
                    if (pendingShortLines.Count > 0)
                    {
                        var joined = string.Join("；", pendingShortLines);
                        if (joined.Length >= 5)
                            rules.Add(new ParsedRule { Category = category, Subcategory = stdSub,
                                Title = stdTitle, ResultText = joined, SourceFile = fileName });
                        pendingShortLines.Clear();
                    }
                    stdTitle = text;
                }
                else if (text.Length >= 15)
                {
                    if (pendingShortLines.Count > 0)
                    {
                        var joined = string.Join("；", pendingShortLines);
                        if (joined.Length >= 5)
                            rules.Add(new ParsedRule { Category = category, Subcategory = stdSub,
                                Title = stdTitle, ResultText = joined, SourceFile = fileName });
                        pendingShortLines.Clear();
                    }
                    rules.Add(new ParsedRule
                    {
                        Category = category,
                        Subcategory = stdSub,
                        Title = stdTitle,
                        ResultText = text,
                        SourceFile = fileName
                    });
                    if (!isHeading) stdTitle = null;
                }
                else if (text.Length >= 3)
                {
                    pendingShortLines.Add(text);
                }
            }

            if (pendingShortLines.Count > 0)
            {
                var joined = string.Join("；", pendingShortLines);
                if (joined.Length >= 5)
                    rules.Add(new ParsedRule { Category = category, Subcategory = stdSub,
                        Title = stdTitle, ResultText = joined, SourceFile = fileName });
            }

            return rules;
        }

        private static List<ParsedRule> ParseXlsx(Stream stream, string fileName, string category, string? subcategory, bool isLegacy)
        {
            var rules = new List<ParsedRule>();
            IWorkbook workbook = isLegacy ? new HSSFWorkbook(stream) : new XSSFWorkbook(stream);

            for (int si = 0; si < workbook.NumberOfSheets; si++)
            {
                var sheet = workbook.GetSheetAt(si);
                if (sheet == null) continue;
                var sheetName = workbook.GetSheetName(si);

                // Detect header row
                var headerRow = sheet.GetRow(sheet.FirstRowNum);
                int startRow = sheet.FirstRowNum;
                bool hasHeader = headerRow != null && headerRow.Cells.Any(c => !string.IsNullOrWhiteSpace(GetCellValue(c)));
                if (hasHeader) startRow++;

                for (int ri = startRow; ri <= sheet.LastRowNum; ri++)
                {
                    var row = sheet.GetRow(ri);
                    if (row == null) continue;

                    var cells = Enumerable.Range(row.FirstCellNum, Math.Max(0, row.LastCellNum - row.FirstCellNum))
                        .Select(ci => GetCellValue(row.GetCell(ci)))
                        .ToList();

                    if (cells.All(c => string.IsNullOrWhiteSpace(c))) continue;

                    string title = cells.Count > 0 ? cells[0] ?? "" : "";
                    string result = cells.Count > 1
                        ? string.Join("；", cells.Skip(1).Where(c => !string.IsNullOrWhiteSpace(c)))
                        : title;
                    if (cells.Count == 1) { result = title; title = ""; }
                    if (string.IsNullOrWhiteSpace(result)) continue;

                    rules.Add(new ParsedRule
                    {
                        Category = category,
                        Subcategory = string.IsNullOrWhiteSpace(subcategory) ? sheetName : subcategory,
                        Title = string.IsNullOrWhiteSpace(title) ? null : title,
                        ResultText = result,
                        SourceFile = fileName
                    });
                }
            }
            return rules;
        }

        // ---- utilities ----

        private static byte[] ReadAllBytes(Stream stream)
        {
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            return ms.ToArray();
        }

        private static Encoding DetectEncoding(byte[] bytes)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            // BOM check
            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
                return Encoding.UTF8;
            if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
                return Encoding.Unicode;
            // Heuristic: try UTF-8 validity
            try
            {
                var decoded = Encoding.UTF8.GetString(bytes);
                if (!decoded.Contains('\uFFFD')) return Encoding.UTF8;
            }
            catch { }
            // Fallback to Big5 (Traditional Chinese)
            try { return Encoding.GetEncoding("big5"); } catch { }
            return Encoding.GetEncoding("gb2312");
        }

        private static List<string> SplitCsvLine(string line)
        {
            var result = new List<string>();
            bool inQuotes = false;
            var current = new StringBuilder();
            foreach (char c in line)
            {
                if (c == '"') { inQuotes = !inQuotes; }
                else if (c == ',' && !inQuotes) { result.Add(current.ToString()); current.Clear(); }
                else { current.Append(c); }
            }
            result.Add(current.ToString());
            return result;
        }

        private static string GetCellValue(ICell? cell)
        {
            if (cell == null) return "";
            return cell.CellType switch
            {
                CellType.String => cell.StringCellValue ?? "",
                CellType.Numeric => cell.NumericCellValue.ToString(CultureInfo.InvariantCulture),
                CellType.Boolean => cell.BooleanCellValue.ToString(),
                CellType.Formula => cell.CachedFormulaResultType == CellType.String
                    ? cell.StringCellValue ?? ""
                    : cell.NumericCellValue.ToString(CultureInfo.InvariantCulture),
                _ => ""
            };
        }
    }

    // DTOs
    public class FortuneRuleDto
    {
        public string Category { get; set; } = string.Empty;
        public string? Subcategory { get; set; }
        public string? Title { get; set; }
        public string? ConditionText { get; set; }
        public string ResultText { get; set; } = string.Empty;
        public string? SourceFile { get; set; }
        public string? Tags { get; set; }
        public bool? IsActive { get; set; }
        public int? SortOrder { get; set; }
    }

    public class ParsedRule
    {
        public string Category { get; set; } = string.Empty;
        public string? Subcategory { get; set; }
        public string? Title { get; set; }
        public string? ConditionText { get; set; }
        public string ResultText { get; set; } = string.Empty;
        public string? SourceFile { get; set; }
        public string? Tags { get; set; }
    }

    public class ImportRequest
    {
        public string FileName { get; set; } = string.Empty;
        public string FileType { get; set; } = string.Empty;
        public string Category { get; set; } = string.Empty;
        public List<ParsedRule> Rules { get; set; } = new();
    }
}
