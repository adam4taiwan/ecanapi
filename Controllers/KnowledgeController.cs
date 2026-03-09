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
        public async Task<IActionResult> Upload([FromForm] IFormFile file, [FromForm] string category, [FromForm] string? subcategory)
        {
            if (!IsAdmin()) return Forbid();
            if (file == null || file.Length == 0) return BadRequest(new { message = "no file" });

            var ext = Path.GetExtension(file.FileName).ToLowerInvariant();
            var allowed = new[] { ".csv", ".txt", ".docx", ".xlsx", ".xls" };
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
                    ".txt" => ParseTxt(ms, file.FileName, category, subcategory),
                    ".docx" => ParseDocx(ms, file.FileName, category, subcategory),
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
            // Try UTF-8 first, fallback to Big5
            Encoding enc;
            try
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                var bytes = ReadAllBytes(stream);
                enc = DetectEncoding(bytes);
                var text = enc.GetString(bytes);
                var lines = text.Split('\n', StringSplitOptions.RemoveEmptyEntries);
                if (lines.Length == 0) return rules;

                // Check if first line is a header
                var firstLine = lines[0].Trim();
                bool hasHeader = !firstLine.StartsWith("　") && lines.Length > 1;
                int start = hasHeader ? 1 : 0;

                for (int i = start; i < lines.Length; i++)
                {
                    var cols = SplitCsvLine(lines[i]);
                    if (cols.Count == 0) continue;

                    string title = cols.Count > 0 ? cols[0].Trim() : "";
                    string result = cols.Count > 1 ? string.Join("；", cols.Skip(1).Select(c => c.Trim()).Where(c => c.Length > 0)) : title;
                    if (cols.Count == 1) { result = title; title = ""; }
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
            var rules = new List<ParsedRule>();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var bytes = ReadAllBytes(stream);
            var enc = DetectEncoding(bytes);
            var text = enc.GetString(bytes);

            // Split by paragraph (double newline or numbered sections)
            var paragraphs = Regex.Split(text, @"\r?\n\r?\n+")
                .Select(p => p.Trim())
                .Where(p => p.Length > 10)
                .ToList();

            if (paragraphs.Count == 0)
            {
                // fallback: split by single newline
                paragraphs = text.Split('\n')
                    .Select(p => p.Trim())
                    .Where(p => p.Length > 10)
                    .ToList();
            }

            foreach (var para in paragraphs)
            {
                var lines = para.Split('\n');
                var title = lines[0].Trim().Length <= 60 ? lines[0].Trim() : null;
                var body = title != null && lines.Length > 1 ? string.Join(" ", lines.Skip(1)) : para;
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
            return rules;
        }

        private static List<ParsedRule> ParseDocx(Stream stream, string fileName, string category, string? subcategory)
        {
            var rules = new List<ParsedRule>();
            using var doc = WordprocessingDocument.Open(stream, false);
            var body = doc.MainDocumentPart?.Document?.Body;
            if (body == null) return rules;

            var paragraphs = body.Elements<Paragraph>().ToList();
            string? currentTitle = null;
            var currentContent = new StringBuilder();

            void flush()
            {
                var content = currentContent.ToString().Trim();
                if (content.Length < 5) return;
                rules.Add(new ParsedRule
                {
                    Category = category,
                    Subcategory = subcategory,
                    Title = currentTitle,
                    ResultText = content,
                    SourceFile = fileName
                });
                currentTitle = null;
                currentContent.Clear();
            }

            foreach (var para in paragraphs)
            {
                var text = para.InnerText.Trim();
                if (string.IsNullOrWhiteSpace(text)) continue;

                // Check if this paragraph is a heading (bold or short title)
                bool isHeading = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value?.StartsWith("Heading") == true;
                bool isShortBold = text.Length <= 50 && para.Descendants<Bold>().Any();

                if (isHeading || isShortBold)
                {
                    flush();
                    currentTitle = text;
                }
                else
                {
                    if (currentContent.Length > 0) currentContent.Append('\n');
                    currentContent.Append(text);
                    // flush if content is long enough to be a standalone rule
                    if (currentContent.Length > 200 && (text.EndsWith("。") || text.EndsWith(".")))
                    {
                        flush();
                    }
                }
            }
            flush();
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
