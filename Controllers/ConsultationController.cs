using Ecanapi.Data;
using Ecanapi.Models;
using Ecanapi.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Stripe;
using Stripe.Checkout;
using System.Security.Claims;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace Ecanapi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ConsultationController : ControllerBase
    {
        private readonly IAstrologyService _astrologyService;
        private readonly ApplicationDbContext _context;
        private readonly CalendarDbContext _calendarDb;
        private readonly IConfiguration _config;
        private readonly ILogger<ConsultationController> _logger;
        private readonly IEmailService _email;
        private static readonly HttpClient _httpClient = new HttpClient { Timeout = TimeSpan.FromMinutes(5) };

        public ConsultationController(IAstrologyService astrologyService, ApplicationDbContext context, CalendarDbContext calendarDb, IConfiguration config, ILogger<ConsultationController> logger, IEmailService email)
        {
            _astrologyService = astrologyService;
            _context = context;
            _calendarDb = calendarDb;
            _config = config;
            _logger = logger;
            _email = email;
        }

        [Authorize]
        [HttpPost("analyze")]
        public async Task<IActionResult> GetAiAnalysis([FromBody] AiRequest request)
        {
            var identity = User.FindFirstValue(ClaimTypes.Email)
                         ?? User.FindFirstValue(ClaimTypes.Name)
                         ?? User.FindFirst("unique_name")?.Value;

            if (string.IsNullOrEmpty(identity))
                return Unauthorized(new { error = "請重新登入" });

            var user = await _context.Users.FirstOrDefaultAsync(u => u.UserName == identity || u.Email == identity);
            if (user == null) return BadRequest(new { error = $"找不到用戶: {identity}" });

            if (request.Type == "同步查詢")
                return Ok(new { });

            // 訂閱驗證：依命書類型對應 ProductCode
            string reqProductCode = request.Type switch
            {
                "綜合性命書" or "綜合鑑定" => "BOOK_BAZI",
                "大運命書" => "BOOK_DAIYUN",
                "流年命書" => "BOOK_LIUNIAN",
                "問事" => "TOPIC_CONSULT",
                _ => "BOOK_BAZI"
            };

            int analyzeSubId;
            if (reqProductCode == "TOPIC_CONSULT")
            {
                var (topicOk, topicErr) = await CheckTopicConsultAccess(user.Id);
                if (!topicOk) return BadRequest(new { error = topicErr });
                analyzeSubId = 0;
            }
            else
            {
                var (quotaOk, quotaErr, quotaSubId) = await CheckSubscriptionQuota(user.Id, reqProductCode);
                if (!quotaOk) return BadRequest(new { error = quotaErr });
                analyzeSubId = quotaSubId;
            }

            try
            {
                var chartData = await _astrologyService.CalculateChartAsync(request.ChartRequest);
                string chartJson = JsonSerializer.Serialize(chartData);
                int nowYear = DateTime.Now.Year;
                int targetYear = request.TargetYear ?? nowYear;
                int birthYear = request.ChartRequest.Year;
                int currentAge = nowYear - birthYear;

                string prompt = request.Type switch
                {
                    "綜合性命書" or "綜合鑑定" => BuildComprehensivePrompt(chartJson),
                    "大運命書" => BuildMajorLuckPrompt(chartJson, chartData, currentAge, nowYear, request.FortuneDuration ?? 5),
                    "流年命書" => BuildAnnualLuckPrompt(chartJson, targetYear, targetYear - birthYear),
                    "問事" => BuildTopicPrompt(chartJson, request.Topic ?? ""),
                    _ => BuildComprehensivePrompt(chartJson)
                };

                string aiResult = await CallGeminiApi(prompt);

                // 九星氣學加成（純 KB，不扣點）
                if (request.Type is "綜合性命書" or "綜合鑑定" or "大運命書" or "流年命書")
                {
                    string nsSection = await NsBuildBirthSection(
                        request.ChartRequest.Year, request.ChartRequest.Month,
                        request.ChartRequest.Day,  request.ChartRequest.Hour,
                        request.ChartRequest.Gender);
                    if (!string.IsNullOrEmpty(nsSection))
                        aiResult += nsSection;
                }

                if (reqProductCode != "TOPIC_CONSULT" && analyzeSubId > 0)
                    await RecordSubscriptionClaim(user.Id, analyzeSubId, reqProductCode);

                // Save report history
                string reportTitle = request.Type switch
                {
                    "綜合性命書" or "綜合鑑定" => "綜合命書",
                    "大運命書" => $"大運命書（{request.FortuneDuration ?? 5}年）",
                    "流年命書" => $"{request.TargetYear ?? DateTime.Now.Year} 流年命書",
                    "問事" => "問事鑑定",
                    _ => "命理鑑定"
                };
                if (reqProductCode != "TOPIC_CONSULT")
                    await SaveUserReportAsync(user.Id, request.Type ?? "comprehensive", reportTitle, aiResult,
                        new { birthYear = request.ChartRequest.Year, birthMonth = request.ChartRequest.Month, birthDay = request.ChartRequest.Day, gender = request.ChartRequest.Gender });

                return Ok(new { result = aiResult });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "命理鑑定失敗 Type={Type} User={User}", request.Type, identity);
                return StatusCode(500, new { error = "命理鑑定失敗，請稍後再試", details = ex.Message });
            }
        }

        private string BuildComprehensivePrompt(string chartJson) => $@"
你是一位精通《三命通會》與《紫微斗數》的命理鑑定大師『玉洞子』。
請根據以下 JSON 命盤數據進行深度解析，並「嚴格」遵守下方的【模版排版規範】。

### 命盤數據：
{chartJson}

### 輸出格式：

一、格局判斷 (依訣評析)
【依據】
    * 八字解析：(描述日主生於何月、干支五行互動，需包含納音五行)
    * 紫微解析：(描述命、財、官、遷之星曜組合)
【格局】
    * 八字：『(格局名稱)』之真格。
    * 紫微：『(格局名稱)』之複合影響。

二、核心特質 (邏輯制化)
【定論】(用一句堅定的話斷死命格)
【解析】(深度解析內在與外在的互動)

三、關鍵斷語 (深度細查)
【財官運勢】：(明確斷言事業成就等級，嚴禁使用「可能」)
【六親情感】：(明確斷言婚姻配偶特徵)
【健康壽元】：(分析身體強弱、注意病符)

四、具體建議 (行為指南)
【行動指南】：(給出白話、具體、可執行的三條動作建議)

-----------------------------------------------------------------
鑑定大師：玉洞子  |  修身齊家，命在人心。
";

        private string BuildMajorLuckPrompt(string chartJson, AstrologyChartResult chartData, int currentAge, int nowYear, int duration)
        {
            // 計算年份範圍
            int endYear = duration == 0 ? nowYear + 80 : nowYear + duration;
            var annualLucks = _astrologyService.GenerateAnnualLucks(nowYear, endYear, chartData.Bazi.DayPillar.HeavenlyStem);
            string annualJson = JsonSerializer.Serialize(annualLucks);

            string durationLabel = duration == 0 ? "終身" : $"{duration}年";
            string detailLevel = duration switch
            {
                5 => "請逐月分析，表格欄位名稱固定為：月份、月柱、六神、吉凶、斷語。每格斷語內容不超過30字，表格標題列只寫欄位名稱。",
                10 => "請逐月分析，表格欄位名稱固定為：月份、月柱、六神、吉凶、斷語。每格斷語內容不超過20字，表格標題列只寫欄位名稱。",
                20 => "請逐季分析（每季3個月合併），表格欄位名稱固定為：季度、干支、六神、吉凶、重點。每格重點內容不超過30字，表格標題列只寫欄位名稱。",
                30 => "請逐年分析，表格欄位名稱固定為：年份、干支、六神、吉凶、重點。每格重點內容不超過30字，表格標題列只寫欄位名稱。",
                _ => "請逐大運分析，表格欄位名稱固定為：大運、干支、重點年份、吉凶、總評。每格總評不超過30字，表格標題列只寫欄位名稱。"
            };

            return $@"
你是一位精通《三命通會》大運推演與《紫微斗數》四化的命理鑑定大師『玉洞子』。
請根據以下命盤數據，從命主當前年齡（{currentAge} 歲，西元 {nowYear} 年）起，
出具【{durationLabel}大運命書】，依每個月吉凶分別判斷。

## 分析依據（務必逐一運用）：
1. 八字干支：合（六合、三合、半三合）、沖（六沖）、刑（三刑）、害（六害）、破（六破）
2. 神煞：天德、月德、天乙、文昌、文曲、羊刃、七煞、白虎、弔客、喪門等
3. 六神（十神）：每月月柱天干六神對日主的生剋制化
4. 紫微斗數：相關宮位四化（化祿/化權/化科/化忌）在大運/流年流月的觸發

## 命盤基礎數據：
{chartJson}

## 流年干支對照表（{nowYear}～{endYear}）：
{annualJson}

## 月柱推算規則（五虎遁年起月）：
甲己年→正月丙寅起；乙庚年→正月戊寅起；丙辛年→正月庚寅起；
丁壬年→正月壬寅起；戊癸年→正月甲寅起；月支：寅卯辰巳午未申酉戌亥子丑

## 輸出要求：
{detailLevel}

## 固定首尾格式：

大運命書 · {durationLabel} · 自 {nowYear} 年起

【大運總覽】
(先列出此期間經歷的大運干支、歲數區間、十神定性)

【逐月/逐季/逐年詳析】
(依上方「輸出要求」格式展開，吉月標▲、凶月標▼、平月標－)

【{durationLabel}行動總綱】
* 最佳出手年份：
* 最需低調年份：
* 財運重點：
* 婚姻重點：
* 事業重點：
* 健康提示：

-----------------------------------------------------------------
鑑定大師：玉洞子  |  知命善用，趨吉避凶。
";
        }

        private string BuildAnnualLuckPrompt(string chartJson, int targetYear, int currentAge) => $@"
你是一位精通流年推演的命理鑑定大師『玉洞子』。
請根據以下 JSON 命盤數據，針對 {targetYear} 年（命主約 {currentAge} 歲）進行深度流年分析。

### 命盤數據：
{chartJson}

### 輸出格式：

一、{targetYear} 年流年總論
【流年天干地支】：(說明該年干支)
【與命盤互動】：(流年干支對本命八字的生剋沖合)
【大運配合】：(此流年所在之大運，吉凶是否呼應)
【年運等級】：(★☆☆☆☆ 至 ★★★★★ + 一句斷語)

二、六大面向逐一研判
【財運】：(明確斷言財富進退，禁用「可能」)
【事業】：(明確斷言職場升遷或波折)
【感情婚姻】：(明確斷言感情動向)
【健康】：(點出需注意的身體部位或病符)
【人際是非】：(官司、小人、合作等)
【重大變動】：(搬遷、轉業、出國等可能性)

三、逐季重點提示
【第一季 (1-3月)】：
【第二季 (4-6月)】：
【第三季 (7-9月)】：
【第四季 (10-12月)】：

四、{targetYear} 年行動建議
【最佳時機】：(宜把握哪幾個月動作)
【謹慎時段】：(哪幾個月需低調)
【具體行動】：(三條可執行建議)

-----------------------------------------------------------------
鑑定大師：玉洞子  |  知運在先，趨吉避凶。
";

        private string BuildTopicPrompt(string chartJson, string topic, string kbFacts = "") => $@"
你是一位精通《三命通會》《滴天髓》《紫微斗數》的命理鑑定大師『玉洞子』。
請根據以下命盤數據，專門針對命主的【{topic}】課題，進行五章式深度命書鑑定。

嚴守以下規則：
1. 每個論斷必須同時引用八字與紫微雙系統佐證，不可只談其中一方
2. 嚴禁使用「可能」「或許」「也許」等模糊詞，必須用明確斷語
3. 時機分析需給出具體年份，不可泛論
4. 建議必須具體可執行，不可空泛
5. 流年吉凶評斷必須與「命理系統預算數據」中的交叉評斷完全一致，不可自行重新推算
{(string.IsNullOrEmpty(kbFacts) ? "" : kbFacts)}
### 完整命盤數據（含八字與紫微）：
{chartJson}

### 問事主題：{topic}

---

## 主題命書：{topic}

### 第一章　命主概況

【八字格局】：(日主五行、身強/身弱判定、格局名稱，一句定位)
【紫微命宮】：(命宮主星及其性質，一句定位)
【核心特質】：(整合八字與紫微，用兩句話說出此命主最鮮明的人格特質)

### 第二章　{topic} 先天格局

【八字論斷】：
  - 哪個干支/用神最直接影響{topic}（需說出具體天干地支）
  - {topic}的先天優勢與先天制約各為何

【紫微論斷】：
  - 與{topic}最相關的宮位（如財帛宮/夫妻宮/官祿宮/田宅宮等）的主星組合
  - 四化飛入此宮的影響（有則說，無則略）

【格局定論】：(八字與紫微雙系統交叉後，給出{topic}整體先天格局的一句斷語)

### 第三章　{topic} 深度研判

【有利因素】：(列出命盤中對{topic}有利的2-3項具體元素)
【不利因素】：(列出命盤中對{topic}的2-3項阻礙)
【輕重比較】：(吉凶相較，{topic}整體偏順還是偏阻？給出明確結論)

### 第四章　時機推算

【當前大運】：(目前走哪個大運，此大運對{topic}的影響)
【近三年流年】：(針對未來3年每年給出具體{topic}走勢預測，格式：「YYYY年：...」)
【最佳行動窗口】：(在可見的大運流年中，哪一年最適合在{topic}上採取重大行動)
【需要謹慎年份】：(哪年不宜在{topic}上輕舉妄動，說明原因)

### 第五章　趨吉行動方案

【立即行動】：(現在就能執行的2-3件具體事項)
【中期佈局】：(未來1-3年的策略方向)
【化煞開運】：(針對{topic}的具體化解建議，含五行補強或擇日方向)

---
鑑定大師：玉洞子  |  問事明心，決策有據。
";

        [HttpPost("create-checkout-session")]
        public async Task<IActionResult> CreateCheckoutSession([FromBody] CreateCheckoutSessionRequest request)
        {
            StripeConfiguration.ApiKey = _config["Stripe:SecretKey"];
            var options = new SessionCreateOptions
            {
                PaymentMethodTypes = new List<string> { "card" },
                LineItems = new List<SessionLineItemOptions> {
                    new SessionLineItemOptions {
                        PriceData = new SessionLineItemPriceDataOptions {
                            UnitAmount = request.Amount,
                            Currency = "twd",
                            ProductData = new SessionLineItemPriceDataProductDataOptions { Name = $"{request.Points} 點數" },
                        },
                        Quantity = 1,
                    },
                },
                Mode = "payment",
                SuccessUrl = request.SuccessUrl + "?session_id={CHECKOUT_SESSION_ID}",
                CancelUrl = request.CancelUrl,
            };
            var service = new SessionService();
            var session = await service.CreateAsync(options);
            return Ok(new { id = session.Id, url = session.Url });
        }

        // ============================================================
        // KB 綜合命書端點（純知識庫，不呼叫 Gemini）
        // ============================================================
        [Authorize]
        [HttpGet("analyze-kb")]
        public async Task<IActionResult> GetKbAnalysis()
        {
            var identity = User.FindFirstValue(ClaimTypes.Email)
                         ?? User.FindFirstValue(ClaimTypes.Name)
                         ?? User.FindFirst("unique_name")?.Value;
            if (string.IsNullOrEmpty(identity))
                return Unauthorized(new { error = "請重新登入" });

            var user = await _context.Users.FirstOrDefaultAsync(u => u.UserName == identity || u.Email == identity);
            if (user == null) return BadRequest(new { error = "找不到用戶" });

            var userChart = await _context.UserCharts.FirstOrDefaultAsync(c => c.UserId == user.Id);
            if (userChart == null || string.IsNullOrEmpty(userChart.ChartJson))
                return BadRequest(new { error = "no_chart" });

            bool kbIsAdmin = string.Equals(user.Email, _config["Admin:Email"], StringComparison.OrdinalIgnoreCase);
            int kbSubId = -1;
            if (!kbIsAdmin)
            {
                var (kbOk, kbErr, kbSubIdVal) = await CheckSubscriptionQuota(user.Id, "BOOK_BAZI");
                if (!kbOk) return BadRequest(new { error = kbErr });
                kbSubId = kbSubIdVal;
            }

            try
            {
                var root = JsonDocument.Parse(userChart.ChartJson).RootElement;

                // === 提取八字四柱 ===
                if (!root.TryGetProperty("bazi", out var bazi) && !root.TryGetProperty("baziInfo", out bazi))
                    return BadRequest(new { error = "命盤資料格式錯誤" });

                var yearP  = bazi.GetProperty("yearPillar");
                var monthP = bazi.GetProperty("monthPillar");
                var dayP   = bazi.GetProperty("dayPillar");
                var timeP  = bazi.GetProperty("timePillar");

                string nianGan = yearP.GetProperty("heavenlyStem").GetString()  ?? "";
                string nianZhi = yearP.GetProperty("earthlyBranch").GetString() ?? "";
                string nianSS  = yearP.GetProperty("heavenlyStemLiuShen").GetString() ?? "";
                string nianNaYin = yearP.GetProperty("naYin").GetString() ?? "";

                string yueGan  = monthP.GetProperty("heavenlyStem").GetString()  ?? "";
                string yueZhi  = monthP.GetProperty("earthlyBranch").GetString() ?? "";
                string yueSS   = monthP.GetProperty("heavenlyStemLiuShen").GetString() ?? "";
                string yueNaYin = monthP.GetProperty("naYin").GetString() ?? "";

                string riGan   = dayP.GetProperty("heavenlyStem").GetString()  ?? "";
                string riZhi   = dayP.GetProperty("earthlyBranch").GetString() ?? "";
                string riNaYin = dayP.GetProperty("naYin").GetString() ?? "";

                string shiGan  = timeP.GetProperty("heavenlyStem").GetString()  ?? "";
                string shiZhi  = timeP.GetProperty("earthlyBranch").GetString() ?? "";
                string shiSS   = timeP.GetProperty("heavenlyStemLiuShen").GetString() ?? "";
                string shiNaYin = timeP.GetProperty("naYin").GetString() ?? "";

                string riZhu  = riGan + riZhi;
                string nianZhu = nianGan + nianZhi;
                string yueZhu  = yueGan + yueZhi;
                string shiZhu  = shiGan + shiZhi;

                // 地支主氣十神
                string nianZhiSS = KbGetFirstHiddenSS(yearP);
                string yueZhiSS  = KbGetFirstHiddenSS(monthP);
                string riZhiSS   = KbGetFirstHiddenSS(dayP);
                string shiZhiSS  = KbGetFirstHiddenSS(timeP);

                // 衍生資訊
                string riWuXing    = KbStemToElement(riGan);
                string season      = KbBranchToSeason(yueZhi);
                string nianShiCombo = nianGan + shiGan;
                string hourCol     = KbBranchToHourCol(shiZhi);
                string nianAnimal  = KbBranchToZodiac(nianZhi);
                string riAnimal    = KbBranchToZodiac(riZhi);

                // 其他欄位：姓名優先用 ChartName（命書用名）
                string userName    = !string.IsNullOrEmpty(user.ChartName) ? user.ChartName
                                   : !string.IsNullOrEmpty(user.Name) ? user.Name
                                   : root.TryGetProperty("userName", out var un) ? un.GetString() ?? "" : "";
                string solarBirth  = root.TryGetProperty("solarBirthDate", out var sb2) ? sb2.GetString() ?? "" : "";
                // 農曆轉中文日期格式
                string lunarRaw    = root.TryGetProperty("lunarBirthDate", out var lb) ? lb.GetString() ?? "" : "";
                string lunarBirth  = KbLunarToChineseStr(lunarRaw);
                string wuXingJu    = root.TryGetProperty("wuXingJuText", out var wx)   ? wx.GetString() ?? "" : "";
                string mingZhu     = root.TryGetProperty("mingZhu", out var mz)   ? mz.GetString() ?? "" : "";
                string shenZhu     = root.TryGetProperty("shenZhu", out var sz)   ? sz.GetString() ?? "" : "";

                string geJu = "", rootType = "", phenomenon = "";
                if (root.TryGetProperty("baziAnalysisResult", out var bar))
                {
                    geJu       = bar.TryGetProperty("shiShen",   out var ss) ? ss.GetString() ?? "" : "";
                    rootType   = bar.TryGetProperty("rootType",  out var rt) ? rt.GetString() ?? "" : "";
                    phenomenon = bar.TryGetProperty("phenomenon",out var ph) ? ph.GetString() ?? "" : "";
                }
                // 解析四柱斷語
                var (pillarHeader, pillarNian, pillarYue, pillarRi, pillarShi) = KbSplitPillars(geJu);

                string mingGongStars = userChart.MingGongMainStars ?? "";

                // 宮位 + 大運
                var palaces = root.TryGetProperty("palaces", out var pArr) ? pArr : default;
                int currentAge = user.BirthYear.HasValue ? DateTime.Today.Year - user.BirthYear.Value : 0;
                var (daYunStem, daYunBranch, daYunSS, daYunStart, daYunEnd) = KbGetCurrentLuck(root, currentAge);
                string daYunPalace = KbGetLuckPalace(palaces, currentAge);

                // === 查詢 DB ===
                // 六十甲子日柱斷語（BaziDayPillarReadings）
                var kbDayPillar = await _context.BaziDayPillarReadings
                    .FirstOrDefaultAsync(r => r.DayPillar == riGan + riZhi);

                // 六十甲子命主（各分析欄位）
                string rgxx  = await KbQuery($"SELECT COALESCE(rgxx,'')  AS \"Value\" FROM public.\"六十甲子命主\" WHERE rgz='{riZhu}'");
                string rgcz  = await KbQuery($"SELECT COALESCE(rgcz,'')  AS \"Value\" FROM public.\"六十甲子命主\" WHERE rgz='{riZhu}'");
                string rgzfx = await KbQuery($"SELECT COALESCE(rgzfx,'') AS \"Value\" FROM public.\"六十甲子命主\" WHERE rgz='{riZhu}'");
                string xgfx  = await KbQuery($"SELECT COALESCE(xgfx,'')  AS \"Value\" FROM public.\"六十甲子命主\" WHERE rgz='{riZhu}'");
                string aqfx  = await KbQuery($"SELECT COALESCE(aqfx,'')  AS \"Value\" FROM public.\"六十甲子命主\" WHERE rgz='{riZhu}'");
                string syfx  = await KbQuery($"SELECT COALESCE(syfx,'')  AS \"Value\" FROM public.\"六十甲子命主\" WHERE rgz='{riZhu}'");
                string cyfx  = await KbQuery($"SELECT COALESCE(cyfx,'')  AS \"Value\" FROM public.\"六十甲子命主\" WHERE rgz='{riZhu}'");
                string jkfx  = await KbQuery($"SELECT COALESCE(jkfx,'')  AS \"Value\" FROM public.\"六十甲子命主\" WHERE rgz='{riZhu}'");

                // 六十甲子納音
                string naYinDesc = await KbQuery($"SELECT COALESCE(\"desc\",'') AS \"Value\" FROM public.\"六十甲子納音\" WHERE \"干支\"='{riZhu}'");

                // 三命論會（日干 + 月份 + 時支）
                string sanMing = await KbQuery($"SELECT COALESCE(\"desc\",'') AS \"Value\" FROM public.\"六十甲子日對時\" WHERE \"Sky\"='{riGan}' AND \"Month\"='{yueZhi}月' AND \"time\" LIKE '%{shiZhi}%'");

                // 五行喜忌（日主五行 + 季節）
                string xiJi = await KbQuery($"SELECT COALESCE(\"sjrs\",'') AS \"Value\" FROM public.\"五行喜忌\" WHERE \"wh\"='{riWuXing}' AND \"sj\"='{season}'");

                // 六神四柱數
                string gdNianGan = await KbQuery($"SELECT COALESCE(\"Desc\",'') AS \"Value\" FROM public.\"六神四柱數\" WHERE \"position\"='年干{nianSS}'");
                string gdYueGan  = await KbQuery($"SELECT COALESCE(\"Desc\",'') AS \"Value\" FROM public.\"六神四柱數\" WHERE \"position\"='月干{yueSS}'");
                string gdNianZhi = await KbQuery($"SELECT COALESCE(\"Desc\",'') AS \"Value\" FROM public.\"六神四柱數\" WHERE \"position\"='年支{nianZhiSS}'");
                string gdYueZhi  = await KbQuery($"SELECT COALESCE(\"Desc\",'') AS \"Value\" FROM public.\"六神四柱數\" WHERE \"position\"='月支{yueZhiSS}'");
                string gdRiZhi   = await KbQuery($"SELECT COALESCE(\"Desc\",'') AS \"Value\" FROM public.\"六神四柱數\" WHERE \"position\"='日支{riZhiSS}'");
                string gdShiGan  = await KbQuery($"SELECT COALESCE(\"Desc\",'') AS \"Value\" FROM public.\"六神四柱數\" WHERE \"position\"='時干{shiSS}'");

                // 十二生肖性向（年支 + 日支）
                string zodiacNian = await KbQuery($"SELECT COALESCE(\"sxgx\",'') AS \"Value\" FROM public.\"十二生肖性向\" WHERE \"sx\" IN ('{nianAnimal}','{nianZhi}') LIMIT 1");
                string zodiacRi   = await KbQuery($"SELECT COALESCE(\"sxgx\",'') AS \"Value\" FROM public.\"十二生肖性向\" WHERE \"sx\" IN ('{riAnimal}','{riZhi}') LIMIT 1");

                // astro_twoheader（年干+時干）
                string astroN    = await KbQuery($"SELECT COALESCE(\"N\",'') AS \"Value\" FROM astro_twoheader WHERE trim(\"A\")='{nianShiCombo}'");
                string astroM    = await KbQuery($"SELECT COALESCE(\"M\",'') AS \"Value\" FROM astro_twoheader WHERE trim(\"A\")='{nianShiCombo}'");
                string astroHour = await KbQuery($"SELECT COALESCE(\"{hourCol}\",'') AS \"Value\" FROM astro_twoheader WHERE trim(\"A\")='{nianShiCombo}'");
                string astroR    = await KbQuery($"SELECT COALESCE(\"R\",'') AS \"Value\" FROM astro_twoheader WHERE trim(\"A\")='{nianShiCombo}'");
                string astroX    = await KbQuery($"SELECT COALESCE(\"X\",'') AS \"Value\" FROM astro_twoheader WHERE trim(\"A\")='{nianShiCombo}'");
                string astroZ    = await KbQuery($"SELECT COALESCE(\"Z\",'') AS \"Value\" FROM astro_twoheader WHERE trim(\"A\")='{nianShiCombo}'");
                string astroT    = await KbQuery($"SELECT COALESCE(\"T\",'') AS \"Value\" FROM astro_twoheader WHERE trim(\"A\")='{nianShiCombo}'");

                // 窮通寶鑑（日干 + 月支）
                string tongBao = await KbQuery($"SELECT COALESCE(content,'') AS \"Value\" FROM public.\"窮通寶鑑\" WHERE tg='{riGan}' AND dz='{yueZhi}'");

                // 年干四化性格（四化干性.docx）
                string nianSiHuaXing = await KbQuery($"SELECT COALESCE(\"ResultText\",'') AS \"Value\" FROM \"FortuneRules\" WHERE \"SourceFile\"='四化干性.docx' AND \"Title\" LIKE '{nianGan}年干%' LIMIT 1");

                // 紫微所在宮位地支（ziwei_patterns_144 查詢基準）
                string ziweiPos     = KbGetZiweiPosition(palaces);

                // ziwei_patterns_144：查一次完整命盤（ziwei_position=紫微地支, palace_position=命宮地支）
                string ziweiFullContent = await KbZiweiFullQuery(palaces, ziweiPos);

                // 從命盤取得所有星曜集合（用於過濾 ziwei 條件段落）
                var chartStars = KbGetAllChartStars(palaces);
                // 加入先天四化組合（如「貪狼化忌」「破軍化祿」），讓「貪狼化權會照」等非當前年干的條件段落被過濾
                KbAddActiveTransformations(chartStars, nianGan);

                // 從完整內容拆出各宮段落，並過濾掉命盤中不存在的星的條件段落
                // 同宮條件用該宮自己的星曜集合；會照/相夾條件用整個命盤星曜集合
                string ziweiMing    = KbFilterZiweiContent(KbExtractPalaceSection(ziweiFullContent, "命宮"), KbGetPalaceStarsSet(palaces, "命宮"), chartStars);
                string ziweiOffStar = KbGetPalaceStars(palaces, "官祿");
                string ziweiOff     = KbFilterZiweiContent(KbExtractPalaceSection(ziweiFullContent, "事業宮"), KbGetPalaceStarsSet(palaces, "官祿"), chartStars);
                string ziweiWltStar = KbGetPalaceStars(palaces, "財帛");
                string ziweiWlt     = KbFilterZiweiContent(KbExtractPalaceSection(ziweiFullContent, "財帛宮"), KbGetPalaceStarsSet(palaces, "財帛"), chartStars);
                string ziweiSpsStar = KbGetPalaceStars(palaces, "夫妻");
                string ziweiSps     = KbFilterZiweiContent(KbExtractPalaceSection(ziweiFullContent, "夫妻宮"), KbGetPalaceStarsSet(palaces, "夫妻"), chartStars);
                string ziweiHltStar = KbGetPalaceStars(palaces, "疾厄");
                string ziweiHlt     = KbFilterZiweiContent(KbExtractPalaceSection(ziweiFullContent, "疾厄宮"), KbGetPalaceStarsSet(palaces, "疾厄"), chartStars);
                string ziweiParStar = KbGetPalaceStars(palaces, "父母");
                string ziweiPar     = KbFilterZiweiContent(KbExtractPalaceSection(ziweiFullContent, "父母宮"), KbGetPalaceStarsSet(palaces, "父母"), chartStars);
                string ziweiCldStar = KbGetPalaceStars(palaces, "子女");
                string ziweiCld     = KbFilterZiweiContent(KbExtractPalaceSection(ziweiFullContent, "子女宮"), KbGetPalaceStarsSet(palaces, "子女"), chartStars);
                // 大運宮位轉為 content 標題格式（官祿宮→事業宮、奴僕宮→交友宮）
                string daYunContentKey = string.IsNullOrEmpty(daYunPalace) ? "" : KbContentPalaceName(daYunPalace);
                string ziweiLuck    = string.IsNullOrEmpty(daYunContentKey) ? "" : KbFilterZiweiContent(KbExtractPalaceSection(ziweiFullContent, daYunContentKey), KbGetPalaceStarsSet(palaces, daYunPalace ?? ""), chartStars);

                // 先天四化（年干→化祿/化權/化科/化忌落宮 + KB 描述）
                string siHuaLuPalace   = KbGetSiHuaPalace(nianGan, "化祿", palaces);
                string siHuaQuanPalace = KbGetSiHuaPalace(nianGan, "化權", palaces);
                string siHuaKePalace   = KbGetSiHuaPalace(nianGan, "化科", palaces);
                string siHuaJiPalace   = KbGetSiHuaPalace(nianGan, "化忌", palaces);
                string siHuaLu   = await KbSiHuaQuery(nianGan, "化祿", palaces);
                string siHuaQuan = await KbSiHuaQuery(nianGan, "化權", palaces);
                string siHuaKe   = await KbSiHuaQuery(nianGan, "化科", palaces);
                string siHuaJi   = await KbSiHuaQuery(nianGan, "化忌", palaces);

                // 宮位四化（各宮宮干→化星飛入目標宮，每宮查全四化）
                var (mingLuPalace,   mingLuContent)   = await KbGongWeiSiHuaQuery(palaces, "命宮",  "化祿");
                var (mingQuanPalace, mingQuanContent) = await KbGongWeiSiHuaQuery(palaces, "命宮",  "化權");
                var (mingKePalace,   mingKeContent)   = await KbGongWeiSiHuaQuery(palaces, "命宮",  "化科");
                var (mingJiPalace,   mingJiContent)   = await KbGongWeiSiHuaQuery(palaces, "命宮",  "化忌");
                var (offLuPalace,    offLuContent)    = await KbGongWeiSiHuaQuery(palaces, "官祿宮","化祿");
                var (offQuanPalace,  offQuanContent)  = await KbGongWeiSiHuaQuery(palaces, "官祿宮","化權");
                var (offKePalace,    offKeContent)    = await KbGongWeiSiHuaQuery(palaces, "官祿宮","化科");
                var (offJiPalace,    offJiContent)    = await KbGongWeiSiHuaQuery(palaces, "官祿宮","化忌");
                var (wltLuPalace,    wltLuContent)    = await KbGongWeiSiHuaQuery(palaces, "財帛宮","化祿");
                var (wltQuanPalace,  wltQuanContent)  = await KbGongWeiSiHuaQuery(palaces, "財帛宮","化權");
                var (wltKePalace,    wltKeContent)    = await KbGongWeiSiHuaQuery(palaces, "財帛宮","化科");
                var (wltJiPalace,    wltJiContent)    = await KbGongWeiSiHuaQuery(palaces, "財帛宮","化忌");
                var (spsLuPalace,    spsLuContent)    = await KbGongWeiSiHuaQuery(palaces, "夫妻宮","化祿");
                var (spsQuanPalace,  spsQuanContent)  = await KbGongWeiSiHuaQuery(palaces, "夫妻宮","化權");
                var (spsKePalace,    spsKeContent)    = await KbGongWeiSiHuaQuery(palaces, "夫妻宮","化科");
                var (spsJiPalace,    spsJiContent)    = await KbGongWeiSiHuaQuery(palaces, "夫妻宮","化忌");
                var (hltLuPalace,    hltLuContent)    = await KbGongWeiSiHuaQuery(palaces, "疾厄宮","化祿");
                var (hltQuanPalace,  hltQuanContent)  = await KbGongWeiSiHuaQuery(palaces, "疾厄宮","化權");
                var (hltKePalace,    hltKeContent)    = await KbGongWeiSiHuaQuery(palaces, "疾厄宮","化科");
                var (hltJiPalace,    hltJiContent)    = await KbGongWeiSiHuaQuery(palaces, "疾厄宮","化忌");

                // 主星入宮（6/7/8三個主星文件）
                string starDescMing = await KbQueryStarInPalace(palaces, "命宮");
                string starDescOff  = await KbQueryStarInPalace(palaces, "官祿宮");
                string starDescWlt  = await KbQueryStarInPalace(palaces, "財帛宮");
                string starDescSps  = await KbQueryStarInPalace(palaces, "夫妻宮");
                string starDescHlt  = await KbQueryStarInPalace(palaces, "疾厄宮");

                // 雙星組合入宮
                string doubleDescMing = await KbQueryDoubleStarInPalace(palaces, "命宮");
                string doubleDescOff  = await KbQueryDoubleStarInPalace(palaces, "官祿宮");
                string doubleDescWlt  = await KbQueryDoubleStarInPalace(palaces, "財帛宮");
                string doubleDescSps  = await KbQueryDoubleStarInPalace(palaces, "夫妻宮");
                string doubleDescHlt  = await KbQueryDoubleStarInPalace(palaces, "疾厄宮");

                // 輔星（吉星/煞星）入宮
                string minorDescMing = await KbQueryMinorStarsInPalace(palaces, "命宮");
                string minorDescOff  = await KbQueryMinorStarsInPalace(palaces, "官祿宮");
                string minorDescWlt  = await KbQueryMinorStarsInPalace(palaces, "財帛宮");
                string minorDescSps  = await KbQueryMinorStarsInPalace(palaces, "夫妻宮");
                string minorDescHlt  = await KbQueryMinorStarsInPalace(palaces, "疾厄宮");

                // 紫微格局偵測 + 描述查詢
                string ziweiGeJu = "";
                if (palaces.ValueKind == JsonValueKind.Array)
                {
                    string mingBr = KbGetPalaceBranch(palaces, "命宮");
                    var gjList = LfDetectZiweiGeJu(mingGongStars, mingBr, chartStars,
                        siHuaLuPalace, siHuaQuanPalace, siHuaKePalace, palaces);
                    var gjSb = new StringBuilder();
                    foreach (var gj in gjList)
                    {
                        string desc = await KbQuery($"SELECT COALESCE(\"ResultText\",'') AS \"Value\" FROM \"FortuneRules\" WHERE \"SourceFile\"='紫微格局說明.docx' AND \"Title\"='{gj}' LIMIT 1");
                        if (!string.IsNullOrEmpty(desc)) { gjSb.AppendLine($"【{gj}】"); gjSb.AppendLine(desc); gjSb.AppendLine(); }
                    }
                    ziweiGeJu = gjSb.ToString();
                }

                // === 組裝命書 ===
                var sb_out = new StringBuilder();

                // --- 年齡適切性提示（高齡者提醒）---
                {
                    string kbAgeHint = LfAgeTopicHint(currentAge);
                    if (!string.IsNullOrEmpty(kbAgeHint)) sb_out.AppendLine(kbAgeHint);
                }

                // --- 一、命盤概覽 ---
                sb_out.AppendLine("=== 一、命盤概覽 ===");
                sb_out.AppendLine($"姓名：{userName}  陽曆：{KbFormatSolarDate(solarBirth)}  農曆：{lunarBirth}");
                sb_out.AppendLine($"四柱：{nianZhu}年 {yueZhu}月 {riZhu}日 {shiZhu}時");
                sb_out.AppendLine($"納音：{nianNaYin} · {yueNaYin} · {riNaYin} · {shiNaYin}");
                sb_out.AppendLine($"日主：{riGan}（{riWuXing}）  五行局：{wuXingJu}");
                sb_out.AppendLine($"命宮主星：{mingGongStars}  命主星：{mingZhu}  身主星：{shenZhu}");
                // 先天四化摘要
                var siHuaLines = new List<string>();
                if (!string.IsNullOrEmpty(siHuaLuPalace))   siHuaLines.Add($"化祿入{siHuaLuPalace}");
                if (!string.IsNullOrEmpty(siHuaQuanPalace)) siHuaLines.Add($"化權入{siHuaQuanPalace}");
                if (!string.IsNullOrEmpty(siHuaKePalace))   siHuaLines.Add($"化科入{siHuaKePalace}");
                if (!string.IsNullOrEmpty(siHuaJiPalace))   siHuaLines.Add($"化忌入{siHuaJiPalace}");
                if (siHuaLines.Count > 0)
                    sb_out.AppendLine($"先天四化：{nianGan}年干 → {string.Join("、", siHuaLines)}");
                sb_out.AppendLine();

                // --- 二、格局用神 ---
                sb_out.AppendLine("=== 二、格局用神 ===");
                sb_out.AppendLine("--- 八字格局論 ---");
                if (!string.IsNullOrEmpty(pillarNian)) sb_out.AppendLine(KbFormatPillarLine(pillarNian, "年柱(根)", "祖先及父母", false));
                if (!string.IsNullOrEmpty(pillarYue))  sb_out.AppendLine(KbFormatPillarLine(pillarYue,  "月柱(苗)", "兄弟姊妹",   false));
                if (!string.IsNullOrEmpty(pillarRi))   sb_out.AppendLine(KbFormatPillarLine(pillarRi,   "日柱(花)", "本人及配偶", false));
                if (!string.IsNullOrEmpty(pillarShi))  sb_out.AppendLine(KbFormatPillarLine(pillarShi,  "時柱(果)", "子女及晚年", true));
                if (!string.IsNullOrEmpty(rootType))   sb_out.AppendLine($"【身強弱根源】{rootType}");
                if (!string.IsNullOrEmpty(phenomenon)) sb_out.AppendLine($"【八字格局綱領】{phenomenon}");
                if (!string.IsNullOrEmpty(xiJi))       sb_out.AppendLine($"【用神喜忌】{xiJi}");
                if (!string.IsNullOrEmpty(tongBao))    sb_out.AppendLine($"【月令精論·窮通寶鑑】{tongBao}");
                if (!string.IsNullOrEmpty(rgzfx))      sb_out.AppendLine($"【日柱綜合論述】{KbStripHtml(rgzfx)}");
                sb_out.AppendLine("--- 紫微格局論 ---");
                if (!string.IsNullOrEmpty(nianSiHuaXing)) sb_out.AppendLine($"【先天特性】{StripNianShengRen(nianSiHuaXing)}");
                if (!string.IsNullOrEmpty(ziweiMing))  sb_out.AppendLine($"【主星宮位】{ziweiMing}");
                if (!string.IsNullOrEmpty(starDescMing)) sb_out.AppendLine($"【命宮星情】{starDescMing}");
                if (!string.IsNullOrEmpty(doubleDescMing)) sb_out.AppendLine($"【命宮雙星論斷】{doubleDescMing}");
                if (!string.IsNullOrEmpty(minorDescMing)) sb_out.AppendLine($"【命宮輔星加臨】{minorDescMing}");
                if (!string.IsNullOrEmpty(ziweiGeJu))
                {
                    sb_out.AppendLine("【命宮格局論斷】");
                    sb_out.AppendLine(ziweiGeJu.TrimEnd());
                }
                // 先天四化與命格關聯
                if (!string.IsNullOrEmpty(siHuaLu))    sb_out.AppendLine($"【先天化祿·{siHuaLuPalace}】{siHuaLu}");
                if (!string.IsNullOrEmpty(siHuaQuan))  sb_out.AppendLine($"【先天化權·{siHuaQuanPalace}】{siHuaQuan}");
                if (!string.IsNullOrEmpty(siHuaJi))    sb_out.AppendLine($"【先天化忌·{siHuaJiPalace}】{siHuaJi}");
                // 命宮宮位四化飛出
                if (!string.IsNullOrEmpty(mingLuContent))   sb_out.AppendLine($"【命宮化祿飛{mingLuPalace}】{mingLuContent}");
                if (!string.IsNullOrEmpty(mingQuanContent)) sb_out.AppendLine($"【命宮化權飛{mingQuanPalace}】{mingQuanContent}");
                if (!string.IsNullOrEmpty(mingKeContent))   sb_out.AppendLine($"【命宮化科飛{mingKePalace}】{mingKeContent}");
                if (!string.IsNullOrEmpty(mingJiContent))   sb_out.AppendLine($"【命宮化忌飛{mingJiPalace}】{mingJiContent}");
                bool ch2BaziHas  = !string.IsNullOrEmpty(phenomenon) || !string.IsNullOrEmpty(rootType);
                bool ch2ZiweiHas = !string.IsNullOrEmpty(ziweiMing);
                if (ch2BaziHas && ch2ZiweiHas) sb_out.AppendLine("【格局交叉驗證】八字與紫微雙重印證，命格論斷可信度高。");
                sb_out.AppendLine();

                // --- 三、性格特質 ---
                sb_out.AppendLine("=== 三、性格特質 ===");
                sb_out.AppendLine("--- 八字性格論 ---");
                if (!string.IsNullOrWhiteSpace(kbDayPillar?.Overview)) sb_out.AppendLine($"【核心】{kbDayPillar.Overview}");
                if (!string.IsNullOrEmpty(rgxx))      sb_out.AppendLine($"【日柱概述】{KbStripHtml(rgxx)}");
                if (!string.IsNullOrEmpty(naYinDesc)) sb_out.AppendLine($"【納音性情·{riNaYin}】{naYinDesc}");
                if (!string.IsNullOrEmpty(rgcz))      sb_out.AppendLine($"【坐星詳解】{KbStripHtml(rgcz)}");
                if (!string.IsNullOrEmpty(xgfx))      sb_out.AppendLine($"【性格分析】{KbStripHtml(xgfx)}");
                KbAppendGd(sb_out, $"年干{nianSS}",   gdNianGan);
                KbAppendGd(sb_out, $"月干{yueSS}",    gdYueGan);
                KbAppendGd(sb_out, $"年支{nianZhiSS}", gdNianZhi);
                KbAppendGd(sb_out, $"月支{yueZhiSS}", gdYueZhi);
                KbAppendGd(sb_out, $"日支{riZhiSS}",  gdRiZhi);
                KbAppendGd(sb_out, $"時干{shiSS}",    gdShiGan);
                if (!string.IsNullOrEmpty(zodiacNian)) sb_out.AppendLine($"【年支{nianZhi}·{nianAnimal}性向】{KbStripHtml(zodiacNian)}");
                if (!string.IsNullOrEmpty(zodiacRi))   sb_out.AppendLine($"【日支{riZhi}·{riAnimal}性向】{KbStripHtml(zodiacRi)}");
                sb_out.AppendLine("--- 紫微性格論 ---");
                // 財帛宮主星對財務個性的影響
                if (!string.IsNullOrEmpty(ziweiWlt))   sb_out.AppendLine($"【財帛宮主星·{ziweiWltStar}（財務個性）】{ziweiWlt}");
                // 官祿宮主星對事業個性的影響
                if (!string.IsNullOrEmpty(ziweiOff))   sb_out.AppendLine($"【官祿宮主星·{ziweiOffStar}（事業個性）】{ziweiOff}");
                if (!string.IsNullOrEmpty(astroN))     sb_out.AppendLine($"【詩評·{astroN}】{KbStripHtml(astroM)}");
                if (!string.IsNullOrEmpty(astroHour))  sb_out.AppendLine($"【先天緣性】{KbStripHtml(astroHour)}");
                bool ch3BaziHas  = !string.IsNullOrEmpty(xgfx) || !string.IsNullOrEmpty(rgxx);
                bool ch3ZiweiHas = !string.IsNullOrEmpty(ziweiWlt) || !string.IsNullOrEmpty(ziweiOff);
                if (ch3BaziHas && ch3ZiweiHas) sb_out.AppendLine("【性格交叉驗證】八字與紫微雙重印證，性格特質論斷可信度高。");
                sb_out.AppendLine();

                // --- 四、事業財運 ---
                sb_out.AppendLine("=== 四、事業財運 ===");
                sb_out.AppendLine("--- 八字事業財運論 ---");
                if (!string.IsNullOrEmpty(xiJi))    sb_out.AppendLine($"【六神喜用分析】{xiJi}");
                if (!string.IsNullOrEmpty(astroR))  sb_out.AppendLine($"【基業發展】{KbStripHtml(astroR)}");
                if (!string.IsNullOrEmpty(syfx))    sb_out.AppendLine($"【事業分析】{KbStripHtml(syfx)}");
                if (!string.IsNullOrEmpty(cyfx))    sb_out.AppendLine($"【財運分析】{KbStripHtml(cyfx)}");
                sb_out.AppendLine("--- 紫微事業財運論 ---");
                // 官祿宮12宮星情 + 宮位四化飛出 + 主星入宮
                if (!string.IsNullOrEmpty(ziweiOff))     sb_out.AppendLine($"【官祿宮·{ziweiOffStar}】{ziweiOff}");
                if (!string.IsNullOrEmpty(starDescOff))  sb_out.AppendLine($"【官祿星性】{starDescOff}");
                if (!string.IsNullOrEmpty(offLuContent))   sb_out.AppendLine($"【官祿宮化祿飛{offLuPalace}】{offLuContent}");
                if (!string.IsNullOrEmpty(offQuanContent)) sb_out.AppendLine($"【官祿宮化權飛{offQuanPalace}】{offQuanContent}");
                if (!string.IsNullOrEmpty(offKeContent))   sb_out.AppendLine($"【官祿宮化科飛{offKePalace}】{offKeContent}");
                if (!string.IsNullOrEmpty(offJiContent))   sb_out.AppendLine($"【官祿宮化忌飛{offJiPalace}】{offJiContent}");
                if (!string.IsNullOrEmpty(doubleDescOff)) sb_out.AppendLine($"【官祿雙星論斷】{doubleDescOff}");
                if (!string.IsNullOrEmpty(minorDescOff)) sb_out.AppendLine($"【官祿輔星加臨】{minorDescOff}");
                // 財帛宮12宮星情 + 宮位四化飛出 + 主星入宮
                if (!string.IsNullOrEmpty(ziweiWlt))     sb_out.AppendLine($"【財帛宮·{ziweiWltStar}】{ziweiWlt}");
                if (!string.IsNullOrEmpty(starDescWlt))  sb_out.AppendLine($"【財帛星性】{starDescWlt}");
                if (!string.IsNullOrEmpty(wltLuContent))   sb_out.AppendLine($"【財帛宮化祿飛{wltLuPalace}】{wltLuContent}");
                if (!string.IsNullOrEmpty(wltQuanContent)) sb_out.AppendLine($"【財帛宮化權飛{wltQuanPalace}】{wltQuanContent}");
                if (!string.IsNullOrEmpty(wltKeContent))   sb_out.AppendLine($"【財帛宮化科飛{wltKePalace}】{wltKeContent}");
                if (!string.IsNullOrEmpty(wltJiContent))   sb_out.AppendLine($"【財帛宮化忌飛{wltJiPalace}】{wltJiContent}");
                if (!string.IsNullOrEmpty(doubleDescWlt)) sb_out.AppendLine($"【財帛雙星論斷】{doubleDescWlt}");
                if (!string.IsNullOrEmpty(minorDescWlt)) sb_out.AppendLine($"【財帛輔星加臨】{minorDescWlt}");
                // 先天四化與事業/財帛關聯
                if (!string.IsNullOrEmpty(siHuaLu) && (siHuaLuPalace == "官祿宮" || siHuaLuPalace == "財帛宮"))
                    sb_out.AppendLine($"【先天化祿加成·{siHuaLuPalace}】{siHuaLu}");
                if (!string.IsNullOrEmpty(siHuaJi) && (siHuaJiPalace == "官祿宮" || siHuaJiPalace == "財帛宮"))
                    sb_out.AppendLine($"【先天化忌警示·{siHuaJiPalace}】{siHuaJi}");
                bool ch4BaziHas  = !string.IsNullOrEmpty(xiJi) || !string.IsNullOrEmpty(syfx) || !string.IsNullOrEmpty(cyfx);
                bool ch4ZiweiHas = !string.IsNullOrEmpty(ziweiOff) || !string.IsNullOrEmpty(ziweiWlt);
                if (ch4BaziHas && ch4ZiweiHas) sb_out.AppendLine("【事業財運交叉驗證】八字與紫微雙重印證，事業財運走向明確。");
                else if (ch4BaziHas && !ch4ZiweiHas) sb_out.AppendLine("【事業財運交叉驗證】八字論據清晰；紫微宮位資料待補充，建議以八字為主。");
                else if (!ch4BaziHas && ch4ZiweiHas) sb_out.AppendLine("【事業財運交叉驗證】紫微論據清晰；八字資料待深入，建議以紫微為參考。");
                sb_out.AppendLine();

                // --- 五、婚姻感情 ---
                if (!LfShouldSkipPalace("夫妻宮", currentAge))
                {
                    sb_out.AppendLine("=== 五、婚姻感情 ===");
                    sb_out.AppendLine("--- 八字婚姻感情論 ---");
                    if (!string.IsNullOrEmpty(aqfx))   sb_out.AppendLine($"【感情特質】{KbStripHtml(aqfx)}");
                    if (!string.IsNullOrEmpty(astroX)) sb_out.AppendLine($"【婚姻論斷】{KbStripHtml(astroX)}");
                    sb_out.AppendLine("--- 紫微婚姻感情論 ---");
                    // 夫妻宮12宮星情 + 宮位四化飛出 + 主星入宮
                    if (!string.IsNullOrEmpty(ziweiSps))     sb_out.AppendLine($"【夫妻宮·{ziweiSpsStar}】{ziweiSps}");
                    if (!string.IsNullOrEmpty(starDescSps))  sb_out.AppendLine($"【夫妻星性】{starDescSps}");
                    if (!string.IsNullOrEmpty(spsLuContent))   sb_out.AppendLine($"【夫妻宮化祿飛{spsLuPalace}】{spsLuContent}");
                    if (!string.IsNullOrEmpty(spsQuanContent)) sb_out.AppendLine($"【夫妻宮化權飛{spsQuanPalace}】{spsQuanContent}");
                    if (!string.IsNullOrEmpty(spsKeContent))   sb_out.AppendLine($"【夫妻宮化科飛{spsKePalace}】{spsKeContent}");
                    if (!string.IsNullOrEmpty(spsJiContent))   sb_out.AppendLine($"【夫妻宮化忌飛{spsJiPalace}】{spsJiContent}");
                    if (!string.IsNullOrEmpty(doubleDescSps)) sb_out.AppendLine($"【夫妻雙星論斷】{doubleDescSps}");
                    if (!string.IsNullOrEmpty(minorDescSps)) sb_out.AppendLine($"【夫妻輔星加臨】{minorDescSps}");
                    // 先天四化入夫妻宮
                    if (!string.IsNullOrEmpty(siHuaLu) && siHuaLuPalace == "夫妻宮")
                        sb_out.AppendLine($"【先天化祿加成·夫妻宮】{siHuaLu}");
                    if (!string.IsNullOrEmpty(siHuaJi) && siHuaJiPalace == "夫妻宮")
                        sb_out.AppendLine($"【先天化忌警示·夫妻宮】{siHuaJi}");
                    bool ch5BaziHas  = !string.IsNullOrEmpty(aqfx) || !string.IsNullOrEmpty(astroX);
                    bool ch5ZiweiHas = !string.IsNullOrEmpty(ziweiSps);
                    if (ch5BaziHas && ch5ZiweiHas) sb_out.AppendLine("【婚姻感情交叉驗證】八字與紫微雙重印證，感情婚姻論斷可信度高。");
                    else if (ch5BaziHas && !ch5ZiweiHas) sb_out.AppendLine("【婚姻感情交叉驗證】八字論據清晰；紫微夫妻宮資料待補充。");
                    sb_out.AppendLine();
                }

                // --- 六、健康壽元 ---
                sb_out.AppendLine("=== 六、健康壽元 ===");
                sb_out.AppendLine("--- 八字健康壽元論 ---");
                if (!string.IsNullOrEmpty(jkfx))   sb_out.AppendLine($"【健康傾向】{KbStripHtml(jkfx)}");
                sb_out.AppendLine("--- 紫微健康壽元論 ---");
                // 疾厄宮12宮星情 + 化忌飛出 + 主星入宮
                if (!string.IsNullOrEmpty(ziweiHlt))     sb_out.AppendLine($"【疾厄宮·{ziweiHltStar}】{ziweiHlt}");
                if (!string.IsNullOrEmpty(starDescHlt))  sb_out.AppendLine($"【疾厄星性】{starDescHlt}");
                if (!string.IsNullOrEmpty(hltLuContent))   sb_out.AppendLine($"【疾厄宮化祿飛{hltLuPalace}】{hltLuContent}");
                if (!string.IsNullOrEmpty(hltQuanContent)) sb_out.AppendLine($"【疾厄宮化權飛{hltQuanPalace}】{hltQuanContent}");
                if (!string.IsNullOrEmpty(hltKeContent))   sb_out.AppendLine($"【疾厄宮化科飛{hltKePalace}】{hltKeContent}");
                if (!string.IsNullOrEmpty(hltJiContent))   sb_out.AppendLine($"【疾厄宮化忌飛{hltJiPalace}】{hltJiContent}");
                if (!string.IsNullOrEmpty(doubleDescHlt)) sb_out.AppendLine($"【疾厄雙星論斷】{doubleDescHlt}");
                if (!string.IsNullOrEmpty(minorDescHlt)) sb_out.AppendLine($"【疾厄輔星加臨】{minorDescHlt}");
                // 先天化忌入疾厄
                if (!string.IsNullOrEmpty(siHuaJi) && siHuaJiPalace == "疾厄宮")
                    sb_out.AppendLine($"【先天化忌警示·疾厄宮】{siHuaJi}");
                bool ch6BaziHas  = !string.IsNullOrEmpty(jkfx);
                bool ch6ZiweiHas = !string.IsNullOrEmpty(ziweiHlt);
                if (ch6BaziHas && ch6ZiweiHas) sb_out.AppendLine("【健康交叉驗證】八字與紫微雙重印證，健康關注點明確。");
                sb_out.AppendLine();

                // --- 七、家庭緣份 ---
                sb_out.AppendLine("=== 七、家庭緣份 ===");
                sb_out.AppendLine("--- 八字家庭緣份論 ---");
                if (!string.IsNullOrEmpty(astroT)) sb_out.AppendLine($"【兄弟緣份】{KbStripHtml(astroT)}");
                // 子息緣份：未成年者不談子女
                if (!LfShouldSkipPalace("子女宮", currentAge) && !string.IsNullOrEmpty(astroZ))
                    sb_out.AppendLine($"【{LfPalaceAgeLabel("子女宮", currentAge)}緣份】{KbStripHtml(astroZ)}");
                sb_out.AppendLine("--- 紫微家庭緣份論 ---");
                // 父母宮：65歲以上改標籤，已在 LfShouldSkipPalace 以 65 為界
                if (!LfShouldSkipPalace("父母宮", currentAge) && !string.IsNullOrEmpty(ziweiPar))
                    sb_out.AppendLine($"【{LfPalaceAgeLabel("父母宮", currentAge)}·{ziweiParStar}】{ziweiPar}");
                if (!LfShouldSkipPalace("子女宮", currentAge) && !string.IsNullOrEmpty(ziweiCld))
                    sb_out.AppendLine($"【{LfPalaceAgeLabel("子女宮", currentAge)}·{ziweiCldStar}】{ziweiCld}");
                if (!string.IsNullOrEmpty(siHuaJi) && (siHuaJiPalace == "父母宮" || siHuaJiPalace == "子女宮"))
                {
                    bool jiFatherOk = siHuaJiPalace == "父母宮" && !LfShouldSkipPalace("父母宮", currentAge);
                    bool jiChildOk  = siHuaJiPalace == "子女宮" && !LfShouldSkipPalace("子女宮", currentAge);
                    if (jiFatherOk || jiChildOk)
                        sb_out.AppendLine($"【先天化忌警示·{siHuaJiPalace}】{siHuaJi}");
                }
                sb_out.AppendLine();

                // 第八章大運流年已移除（綜合命書不含大運，另由大運命書提供）
                sb_out.AppendLine("-----------------------------------------------------------------");
                sb_out.AppendLine("命理鑑定大師：玉洞子  |  修身齊家，命在人心。  v3.0");

                if (!kbIsAdmin) await RecordSubscriptionClaim(user.Id, kbSubId, "BOOK_BAZI");
                await SaveUserReportAsync(user.Id, "bazi", "八字命書", sb_out.ToString(),
                    new { birthYear = user.BirthYear, birthMonth = user.BirthMonth, birthDay = user.BirthDay, gender = user.BirthGender });
                return Ok(new { result = sb_out.ToString() });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "KB命理鑑定失敗 User={User}", identity);
                return StatusCode(500, new { error = "命理鑑定失敗，請稍後再試", details = ex.Message });
            }
        }

        // --- KB Helper: 查詢第一筆純文字 ---
        private async Task<string> KbQuery(string sql)
        {
            try
            {
                var result = await _context.Database.SqlQueryRaw<string>(sql).FirstOrDefaultAsync();
                return result ?? "";
            }
            catch { return ""; }
        }

        // --- KB Helper: 取宮位主星縮寫 ---
        private static string KbGetPalaceStars(JsonElement palaces, string palaceName)
        {
            if (palaces.ValueKind != JsonValueKind.Array) return "";
            foreach (var p in palaces.EnumerateArray())
            {
                string pname = p.TryGetProperty("palaceName", out var pn) ? pn.GetString() ?? "" :
                               p.TryGetProperty("name", out var n2) ? n2.GetString() ?? "" : "";
                if (!KbPalaceSame(pname, palaceName)) continue;
                string majorKey = p.TryGetProperty("majorStars", out _) ? "majorStars" : "mainStars";
                if (p.TryGetProperty(majorKey, out var stars) && stars.ValueKind == JsonValueKind.Array)
                    return string.Join(",", stars.EnumerateArray().Select(s => s.GetString() ?? "").Where(s => s.Length > 0));
            }
            return "";
        }

        // --- KB Helper: 查 ziwei_patterns_144 完整命盤（palace_position=命宮地支）---
        private async Task<string> KbZiweiFullQuery(JsonElement palaces, string ziweiPos)
        {
            if (string.IsNullOrEmpty(ziweiPos)) return "";
            string mingGongBranch = KbGetPalaceBranch(palaces, "命宮");
            if (string.IsNullOrEmpty(mingGongBranch)) return "";
            return await KbQuery($"SELECT COALESCE(content,'') AS \"Value\" FROM public.ziwei_patterns_144 WHERE ziwei_position='{ziweiPos}' AND palace_position='{mingGongBranch}' LIMIT 1");
        }

        // --- KB Helper: 以指定地支為命宮查詢紫微星盤格局（大限命宮用）---
        private async Task<string> KbZiweiQueryByBranch(string ziweiPos, string palaceBranch)
        {
            if (string.IsNullOrEmpty(ziweiPos) || string.IsNullOrEmpty(palaceBranch)) return "";
            return await KbQuery($"SELECT COALESCE(content,'') AS \"Value\" FROM public.ziwei_patterns_144 WHERE ziwei_position='{ziweiPos}' AND palace_position='{palaceBranch}' LIMIT 1");
        }

        // --- KB Helper: 從完整命盤內容拆出指定宮位段落 ---
        private static string KbExtractPalaceSection(string fullContent, string palaceName)
        {
            if (string.IsNullOrEmpty(fullContent) || string.IsNullOrEmpty(palaceName)) return "";
            var lines = fullContent.Split('\n');
            int startIdx = -1;
            for (int i = 0; i < lines.Length; i++)
            {
                var line = lines[i].TrimStart();
                if (line.StartsWith(palaceName) &&
                    (line.Contains("在(") || line.Contains("主星") || line.Contains("無主星")))
                {
                    startIdx = i;
                    break;
                }
            }
            if (startIdx < 0) return "";
            var palaceHeaders = new[] { "命宮", "父母宮", "福德宮", "田宅宮", "事業宮", "交友宮", "遷移宮", "疾厄宮", "財帛宮", "子女宮", "夫妻宮", "兄弟宮" };
            int endIdx = lines.Length;
            for (int i = startIdx + 1; i < lines.Length; i++)
            {
                var line = lines[i].TrimStart();
                if (palaceHeaders.Any(p => p != palaceName && line.StartsWith(p) &&
                    (line.Contains("在(") || line.Contains("主星") || line.Contains("無主星"))))
                {
                    endIdx = i;
                    break;
                }
            }
            string extracted = string.Join("\n", lines.Skip(startIdx).Take(endIdx - startIdx)).Trim();
            return KbRemoveOtherPalaceLines(extracted, palaceName);
        }

        // --- KB Helper: 過濾宮位 section 中偷帶其他宮位論述的行 ---
        // 例：夫妻宮 section 內出現「事業宮見陀羅」「紫微在事業宮」等，予以過濾
        private static string KbRemoveOtherPalaceLines(string extracted, string targetPalace)
        {
            if (string.IsNullOrEmpty(extracted)) return extracted;
            var allPalaces = new[] { "命宮", "父母宮", "福德宮", "田宅宮", "事業宮", "交友宮", "遷移宮", "疾厄宮", "財帛宮", "子女宮", "夫妻宮", "兄弟宮" };
            var otherPalaces = allPalaces.Where(p => p != targetPalace).ToArray();
            string targetShort = targetPalace.TrimEnd('宮'); // 夫妻宮→夫妻

            var lines = extracted.Split('\n');
            var result = new List<string>();
            foreach (var line in lines)
            {
                var trimmed = line.TrimEnd();
                // 正規化：移除空白字元（含全形空格），以識別「命　宮」等帶空格寫法
                string normalized = System.Text.RegularExpressions.Regex.Replace(trimmed, @"\s+", "");
                bool mentionsOther = otherPalaces.Any(p => normalized.Contains(p));
                bool mentionsTarget = normalized.Contains(targetPalace) || normalized.Contains(targetShort);
                // 只過濾：提到其他宮位 且 完全沒提目標宮位的行
                if (mentionsOther && !mentionsTarget)
                    continue;
                result.Add(trimmed);
            }
            return string.Join("\n", result).TrimEnd();
        }

        // --- KB Helper: 宮位名正規化（去宮/去身 → XXX宮格式）---
        // CCB 儲存: "命宮","命身","官祿","官祿身","父母" 等，統一轉為 "XXX宮"
        private static string KbNormalizePalaceName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "";
            string b = name.TrimEnd('宮').Replace("身", "");
            return string.IsNullOrEmpty(b) ? "" : b + "宮";
        }

        // --- KB Helper: 判斷兩宮位名是否同一宮 ---
        private static bool KbPalaceSame(string a, string b)
            => !string.IsNullOrEmpty(a) && !string.IsNullOrEmpty(b)
               && KbNormalizePalaceName(a) == KbNormalizePalaceName(b);

        // --- KB Helper: 轉換為 ziwei_patterns_144 內容標題格式 ---
        // 官祿宮在 content 內寫為「事業宮」，奴僕宮寫為「交友宮」
        private static string KbContentPalaceName(string palaceName)
        {
            string n = KbNormalizePalaceName(palaceName);
            return n switch { "官祿宮" => "事業宮", "奴僕宮" => "交友宮", _ => n };
        }

        // --- KB Helper: 取宮位地支（EarthlyBranch 只取首字元，格式 "子 北  方" → "子"）---
        private static readonly string[] BranchChars = {"子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥"};

        // 從 earthlyBranch 字串安全提取地支（容錯 "丑 北東北" 和 "丁丑" 兩種格式）
        private static string ExtractBranchChar(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return "";
            string firstPart = raw.Split(' ')[0];
            return BranchChars.Contains(firstPart)
                ? firstPart
                : BranchChars.FirstOrDefault(b => raw.Contains(b)) ?? "";
        }

        private static string KbGetPalaceBranch(JsonElement palaces, string palaceName)
        {
            if (palaces.ValueKind != JsonValueKind.Array) return "";
            foreach (var p in palaces.EnumerateArray())
            {
                string pname = p.TryGetProperty("palaceName", out var pn) ? pn.GetString() ?? "" :
                               p.TryGetProperty("name", out var n2) ? n2.GetString() ?? "" : "";
                if (!KbPalaceSame(pname, palaceName)) continue;
                string raw = p.TryGetProperty("earthlyBranch", out var br) ? br.GetString() ?? "" : "";
                return ExtractBranchChar(raw);
            }
            return "";
        }

        // --- KB Helper: 取紫微所在宮位地支 ---
        // --- 地支順時針環（index=0 子 ... 11 亥），逆時針偏移用於大限宮位展開 ---
        private static readonly string[] DyBranchRing =
            { "子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥" };

        // 從起始地支出發，逆時針偏移 offset 步（用於大限十二宮展開）
        private static string DyGetCCWBranch(string startBranch, int offset)
        {
            int idx = Array.IndexOf(DyBranchRing, startBranch);
            if (idx < 0) return startBranch;
            return DyBranchRing[(idx - offset + 12) % 12];
        }

        // 依地支找命盤宮位名稱
        private static string KbGetPalaceNameByBranch(JsonElement palaces, string branch)
        {
            if (palaces.ValueKind != JsonValueKind.Array) return "";
            foreach (var p in palaces.EnumerateArray())
            {
                string eb = p.TryGetProperty("earthlyBranch", out var br) ? br.GetString() ?? "" : "";
                if (!string.IsNullOrEmpty(eb)) eb = eb.Split(' ')[0];
                if (eb != branch) continue;
                string pname = p.TryGetProperty("palaceName", out var pn) ? pn.GetString() ?? "" :
                               p.TryGetProperty("name",       out var n2) ? n2.GetString() ?? "" : "";
                return KbNormalizePalaceName(pname);
            }
            return "";
        }

        // 依宮位名稱找地支（KbGetPalaceNameByBranch 的反向）
        private static string KbGetBranchByPalaceName(JsonElement palaces, string palaceName)
        {
            if (palaces.ValueKind != JsonValueKind.Array) return "";
            string normTarget = KbNormalizePalaceName(palaceName);
            foreach (var p in palaces.EnumerateArray())
            {
                string pname = p.TryGetProperty("palaceName", out var pn) ? pn.GetString() ?? "" :
                               p.TryGetProperty("name",       out var n2) ? n2.GetString() ?? "" : "";
                if (KbNormalizePalaceName(pname) != normTarget) continue;
                string eb = p.TryGetProperty("earthlyBranch", out var br) ? br.GetString() ?? "" : "";
                return eb.Length > 0 ? eb.Split(' ')[0] : "";
            }
            return "";
        }

        // 依年齡取紫微大限命宮地支（找 decadeAgeRange 包含該歲的宮位）
        private static string DyGetDecadeMingBranch(JsonElement palaces, int age)
        {
            if (palaces.ValueKind != JsonValueKind.Array) return "";
            foreach (var p in palaces.EnumerateArray())
            {
                string range = p.TryGetProperty("decadeAgeRange", out var dr) ? dr.GetString() ?? "" : "";
                var parts = range.Split('-');
                if (parts.Length == 2 && int.TryParse(parts[0], out int ps) && int.TryParse(parts[1], out int pe)
                    && age >= ps && age <= pe)
                {
                    string eb = p.TryGetProperty("earthlyBranch", out var br) ? br.GetString() ?? "" : "";
                    return eb.Length > 0 ? eb.Split(' ')[0] : "";
                }
            }
            return "";
        }

        // 將本命宮位地支轉換為大限宮位名稱（逆時針偏移，命宮=0...父母宮=11）
        private static string DyGetDecadePalaceName(string starBranch, string decadeMingBranch)
        {
            if (string.IsNullOrEmpty(starBranch) || string.IsNullOrEmpty(decadeMingBranch)) return "";
            string[] palaceOrder = { "命宮","兄弟宮","夫妻宮","子女宮","財帛宮","疾厄宮","遷移宮","交友宮","官祿宮","田宅宮","福德宮","父母宮" };
            int mingIdx = Array.IndexOf(DyBranchRing, decadeMingBranch);
            int starIdx = Array.IndexOf(DyBranchRing, starBranch);
            if (mingIdx < 0 || starIdx < 0) return "";
            int offset = (mingIdx - starIdx + 12) % 12;
            return palaceOrder[offset];
        }

        // 流年四化入大限宮位的白話說明（星曜縮寫 + 四化類型 + 大限宮位）
        private static string DyGetDecadeSiHuaNote(string starAbbr, string siHuaType, string decadePalace)
        {
            // 星曜縮寫 → 特性關鍵字
            var starTrait = new Dictionary<string, string>
            {
                {"同","口福享樂、情緒平和"},{"機","謀略變動、學習貴人"},{"陽","名聲地位、男性貴人"},
                {"陰","財庫積蓄、陰柔桃花"},{"廉","情慾桃花、競爭官非"},{"武","財星金錢、鐵腕意志"},
                {"府","財庫穩定、保守守成"},{"紫","帝王領導、貴氣尊榮"},{"貪","才藝欲望、交際桃花"},
                {"巨","口才學問、是非口舌"},{"相","文書印信、輔佐助力"},{"梁","蔭星解厄、醫療長輩"},
                {"殺","衝動改革、一刀兩斷"},{"破","破壞更新、開創求變"},{"昌","文藝才華、考試文書"},
                {"曲","才藝文書、異性緣"},{"魁","貴人提攜"},{"鉞","貴人助力"},
                {"輔","輔佐協助"},{"弼","幕後助力"},{"祿","財祿增益"},{"馬","奔波變動"}
            };
            // 四化 × 大限宮位 基礎說明
            var luNote = new Dictionary<string, string>
            {
                {"命宮","此大限個人魅力旺，健康提升，諸事順遂"},
                {"兄弟宮","此大限兄弟姊妹有助力，人際財源廣"},
                {"夫妻宮","此大限感情婚姻順利，配偶有助益"},
                {"子女宮","此大限子女有成，下屬助力，創意發展佳"},
                {"財帛宮","此大限財源廣進，財運亨通"},
                {"疾厄宮","此大限健康好轉，抵抗力增強"},
                {"遷移宮","此大限出外發展順遂，異地有貴人口福"},
                {"交友宮","此大限人際關係佳，朋友助力大"},
                {"官祿宮","此大限事業有進展，升遷機會多"},
                {"田宅宮","此大限不動產有益，財庫充裕"},
                {"福德宮","此大限心情愉快，享樂休閒機會多"},
                {"父母宮","此大限長輩庇蔭，文書合約順利"},
            };
            var quanNote = new Dictionary<string, string>
            {
                {"命宮","此大限個性強勢，主導力強，宜積極開創"},
                {"兄弟宮","此大限兄弟姊妹有實力，競爭或互助皆有"},
                {"夫妻宮","此大限另一半強勢主導，感情需多溝通"},
                {"子女宮","此大限子女或下屬有能力，但易有主導權爭議"},
                {"財帛宮","此大限財務掌控力強，積極理財有成"},
                {"疾厄宮","此大限體力充沛，意志力強"},
                {"遷移宮","此大限在外有話語權，出外發展有競爭力"},
                {"交友宮","此大限朋友圈有影響力人士，社交活躍"},
                {"官祿宮","此大限事業掌握主導，升遷或創業有力"},
                {"田宅宮","此大限積極置產，家庭主導力強"},
                {"福德宮","此大限意志堅定，享受追求目標的過程"},
                {"父母宮","此大限與長輩有主導性互動，文書事宜需積極處理"},
            };
            var keNote = new Dictionary<string, string>
            {
                {"命宮","此大限名聲口碑佳，有雅士風範"},
                {"兄弟宮","此大限兄弟姊妹或同輩有好口碑"},
                {"夫妻宮","此大限另一半有才名，感情平穩"},
                {"子女宮","此大限子女或下屬有才氣"},
                {"財帛宮","此大限財務名聲佳，理財受人肯定"},
                {"疾厄宮","此大限健康穩定，養生有成"},
                {"遷移宮","此大限在外口碑好，旅遊交流有收穫"},
                {"交友宮","此大限交友廣闊，口碑助力多"},
                {"官祿宮","此大限職場名聲佳，才能獲肯定"},
                {"田宅宮","此大限不動產穩健，家中有文雅氣氛"},
                {"福德宮","此大限精神生活豐富，興趣才藝有發展"},
                {"父母宮","此大限長輩緣佳，學業文書有好結果"},
            };
            var jiNote = new Dictionary<string, string>
            {
                {"命宮","此大限身心壓力較大，健康宜多留意"},
                {"兄弟宮","此大限兄弟姊妹或同輩關係有摩擦，錢財勿借貸"},
                {"夫妻宮","此大限感情婚姻有波折，需多溝通包容"},
                {"子女宮","此大限子女或下屬緣薄，創業投資宜謹慎"},
                {"財帛宮","此大限財運受阻，嚴防破財、詐騙"},
                {"疾厄宮","此大限健康需特別注意，宜定期檢查"},
                {"遷移宮","此大限出外多阻礙，旅途不順，宜減少長途行程"},
                {"交友宮","此大限小人是非多，慎選朋友"},
                {"官祿宮","此大限事業不順，職場是非多，宜守成"},
                {"田宅宮","此大限財庫受損，嚴防不動產糾紛、資產流失"},
                {"福德宮","此大限精神壓力大，情緒起伏，宜修身養性"},
                {"父母宮","此大限長輩緣薄，文書合約需謹慎"},
            };

            string normPalace = KbNormalizePalaceName(decadePalace);
            string baseNote = siHuaType switch
            {
                "化祿" => luNote.GetValueOrDefault(normPalace, ""),
                "化權" => quanNote.GetValueOrDefault(normPalace, ""),
                "化科" => keNote.GetValueOrDefault(normPalace, ""),
                "化忌" => jiNote.GetValueOrDefault(normPalace, ""),
                _ => ""
            };
            if (string.IsNullOrEmpty(baseNote)) return "";

            // 加入星曜特性點綴（化祿/化忌 才加，化權/化科 已有方向性）
            if ((siHuaType == "化祿" || siHuaType == "化忌") && starTrait.TryGetValue(starAbbr, out var trait))
                baseNote += $"，{(siHuaType == "化祿" ? "帶" : "化")}入{(siHuaType == "化祿" ? "口福" : "壓力")}以{trait.Split('、')[0]}為主";

            return baseNote;
        }

        private static string KbGetZiweiPosition(JsonElement palaces)
        {
            if (palaces.ValueKind != JsonValueKind.Array) return "";
            foreach (var p in palaces.EnumerateArray())
            {
                string majorKey = p.TryGetProperty("majorStars", out _) ? "majorStars" : "mainStars";
                if (!p.TryGetProperty(majorKey, out var stars) || stars.ValueKind != JsonValueKind.Array) continue;
                if (!stars.EnumerateArray().Any(s => { var n = s.GetString() ?? ""; return n == "紫" || n == "紫微"; })) continue;
                string raw = p.TryGetProperty("earthlyBranch", out var br) ? br.GetString() ?? "" : "";
                return raw.Length > 0 ? raw.Split(' ')[0] : "";
            }
            return "";
        }

        // --- KB Helper: 副星全名集合（供 ziwei 內容過濾用）---
        private static readonly HashSet<string> KnownSecondaryStarNames = new()
        {
            "左輔","右弼","文昌","文曲","擎羊","陀羅","火星","鈴星",
            "天馬","天魁","天鉞","祿存","地劫","地空","天刑","天姚",
            "紅鸞","天喜","孤辰","寡宿","天虛","天哭","天壽","天福",
            "天官","天巫","破碎","大耗","解神","三臺","八座","恩光"
        };

        // --- KB Helper: 格局名 → 觸發星對照 ---
        private static readonly Dictionary<string, List<string>> FormationStarMap = new()
        {
            {"左右夾命格", new(){"左輔","右弼"}},
            {"文星拱命格", new(){"文昌","文曲"}},
            {"擎羊入廟格", new(){"擎羊"}},
            {"祿馬交馳格", new(){"天馬","祿存"}},
            {"馬頭帶箭格", new(){"擎羊","天馬"}},
        };

        // --- KB Helper: 所有已知副星的 JSON 單字宇宙 ---
        // JSON 中副星以單字儲存（火星=火, 鈴星=鈴, 陀羅=陀, 左輔=左, 文昌=昌 等）
        // 掃描 KB 句子時，遇到此集合中的字，即認定為副星引用，需查宮位是否存在
        private static readonly HashSet<char> AllSecondaryStarSingleChars = new()
        {
            // 六吉星
            '左','右','昌','曲','魁','鉞',
            // 六煞星
            '火','鈴','羊','陀','空','劫',
            // 祿馬雜曜
            '馬','祿','刑','姚','鸞','喜',
            '孤','寡','虛','哭','壽','福',
            '官','巫','碎','耗','神','臺','座','光',
        };

        // --- KB Helper: 三方四正分組（地支 → 三合三方 + 四正對沖）---
        private static readonly Dictionary<string, string[]> SanFangGroups = new()
        {
            {"亥", new[]{"亥","卯","未","巳"}}, {"卯", new[]{"卯","亥","未","酉"}}, {"未", new[]{"未","亥","卯","丑"}},
            {"寅", new[]{"寅","午","戌","申"}}, {"午", new[]{"午","寅","戌","子"}}, {"戌", new[]{"戌","寅","午","辰"}},
            {"巳", new[]{"巳","酉","丑","亥"}}, {"酉", new[]{"酉","巳","丑","卯"}}, {"丑", new[]{"丑","巳","酉","未"}},
            {"申", new[]{"申","子","辰","寅"}}, {"子", new[]{"子","申","辰","午"}}, {"辰", new[]{"辰","申","子","戌"}},
        };

        // --- KB Helper: 取宮位三方四正內的所有星（用於「會照」過濾，比全盤更精確）---
        private static HashSet<string> KbGetSanFangStars(JsonElement palaces, string palaceName)
        {
            if (palaces.ValueKind != JsonValueKind.Array) return KbGetAllChartStars(palaces);
            string palBranch = KbGetPalaceBranch(palaces, palaceName);
            if (string.IsNullOrEmpty(palBranch) || !SanFangGroups.TryGetValue(palBranch, out var sfBranches))
                return KbGetAllChartStars(palaces); // 無法計算時 fallback 至全盤
            var sfBranchSet = new HashSet<string>(sfBranches);
            var stars = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var majorMap = new Dictionary<string, string>
            {
                {"紫","紫微"},{"機","天機"},{"陽","太陽"},{"武","武曲"},
                {"同","天同"},{"廉","廉貞"},{"府","天府"},{"陰","太陰"},
                {"貪","貪狼"},{"巨","巨門"},{"相","天相"},{"梁","天梁"},
                {"殺","七殺"},{"破","破軍"}
            };
            foreach (var p in palaces.EnumerateArray())
            {
                string raw = p.TryGetProperty("earthlyBranch", out var br) ? br.GetString() ?? "" : "";
                string branch = ExtractBranchChar(raw);
                if (!sfBranchSet.Contains(branch)) continue;
                foreach (var key in new[]{"majorStars","mainStars","secondaryStars","goodStars","badStars","smallStars"})
                {
                    if (!p.TryGetProperty(key, out var arr) || arr.ValueKind != JsonValueKind.Array) continue;
                    foreach (var s in arr.EnumerateArray())
                    {
                        var name = s.GetString() ?? "";
                        if (string.IsNullOrEmpty(name)) continue;
                        stars.Add(name);
                        if (majorMap.TryGetValue(name, out var full)) stars.Add(full);
                        foreach (var kn in KnownSecondaryStarNames)
                            if (kn.StartsWith(name) || kn.EndsWith(name))
                                stars.Add(kn);
                    }
                }
            }
            return stars;
        }

        // --- KB Helper: 從命盤 JSON 取得所有星曜名稱（主星+副星）---
        private static HashSet<string> KbGetAllChartStars(JsonElement palaces)
        {
            var stars = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (palaces.ValueKind != JsonValueKind.Array) return stars;
            // 主星縮寫→全名
            var majorMap = new Dictionary<string, string>
            {
                {"紫","紫微"},{"機","天機"},{"陽","太陽"},{"武","武曲"},
                {"同","天同"},{"廉","廉貞"},{"府","天府"},{"陰","太陰"},
                {"貪","貪狼"},{"巨","巨門"},{"相","天相"},{"梁","天梁"},
                {"殺","七殺"},{"破","破軍"}
            };
            foreach (var p in palaces.EnumerateArray())
            {
                foreach (var key in new[]{"majorStars","mainStars","secondaryStars","goodStars","badStars","smallStars"})
                {
                    if (!p.TryGetProperty(key, out var arr) || arr.ValueKind != JsonValueKind.Array) continue;
                    foreach (var s in arr.EnumerateArray())
                    {
                        var name = s.GetString() ?? "";
                        if (string.IsNullOrEmpty(name)) continue;
                        stars.Add(name);
                        if (majorMap.TryGetValue(name, out var full)) stars.Add(full);
                        // 副星縮寫展開：名稱是已知副星全名的前綴或後綴（如"羊"→"擎羊"，"昌"→"文昌"）
                        foreach (var kn in KnownSecondaryStarNames)
                            if (kn.StartsWith(name) || kn.EndsWith(name))
                                stars.Add(kn);
                    }
                }
            }
            return stars;
        }

        // --- KB Helper: 取特定宮位的完整星曜集合（含縮寫展開）---
        private static HashSet<string> KbGetPalaceStarsSet(JsonElement palaces, string palaceName)
        {
            var stars = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (palaces.ValueKind != JsonValueKind.Array || string.IsNullOrEmpty(palaceName)) return stars;
            var majorMap = new Dictionary<string, string>
            {
                {"紫","紫微"},{"機","天機"},{"陽","太陽"},{"武","武曲"},
                {"同","天同"},{"廉","廉貞"},{"府","天府"},{"陰","太陰"},
                {"貪","貪狼"},{"巨","巨門"},{"相","天相"},{"梁","天梁"},
                {"殺","七殺"},{"破","破軍"}
            };
            foreach (var p in palaces.EnumerateArray())
            {
                string pname = p.TryGetProperty("palaceName", out var pn) ? pn.GetString() ?? "" :
                               p.TryGetProperty("name", out var n2) ? n2.GetString() ?? "" : "";
                if (!KbPalaceSame(pname, palaceName)) continue;
                foreach (var key in new[]{"majorStars","mainStars","secondaryStars","goodStars","badStars","smallStars"})
                {
                    if (!p.TryGetProperty(key, out var arr) || arr.ValueKind != JsonValueKind.Array) continue;
                    foreach (var s in arr.EnumerateArray())
                    {
                        var name = s.GetString() ?? "";
                        if (string.IsNullOrEmpty(name)) continue;
                        stars.Add(name);
                        if (majorMap.TryGetValue(name, out var full)) stars.Add(full);
                        foreach (var kn in KnownSecondaryStarNames)
                            if (kn.StartsWith(name) || kn.EndsWith(name))
                                stars.Add(kn);
                    }
                }
                break; // found matching palace
            }
            return stars;
        }

        // 移除 KB 內容中主題不符的行
        // keepIfContains：即使含 offTopicKeywords，若同時含此字則保留（避免誤刪交叉論述）
        private static string KbRemoveOffTopicLines(string content, string[] offTopicKeywords, string[] keepIfContains = null)
        {
            if (string.IsNullOrEmpty(content)) return content;
            var result = new List<string>();
            foreach (var line in content.Split('\n'))
            {
                bool hasOffTopic = offTopicKeywords.Any(k => line.Contains(k));
                if (!hasOffTopic)
                {
                    result.Add(line);
                    continue;
                }
                // 含 off-topic 關鍵字，但也含 keep 關鍵字 → 保留
                if (keepIfContains != null && keepIfContains.Any(k => line.Contains(k)))
                    result.Add(line);
                // 否則過濾掉
            }
            return string.Join("\n", result).Trim();
        }

        // --- KB Helper: 過濾 ziwei 宮位內容，只保留命盤實際有的星才顯示的段落 ---
        // palaceStars: 該宮位自己的星（用於「同宮」類型檢查）
        // allChartStars: 整個命盤的星（用於「會照」「相夾」等跨宮檢查）
        private static string KbFilterZiweiContent(string content, HashSet<string> palaceStars, HashSet<string> allChartStars)
        {
            if (string.IsNullOrEmpty(content)) return content;
            content = content.Replace("[中州派]", ""); // 移除內部門派標籤，保留內容
            var lines = content.Split('\n');
            var result = new StringBuilder();
            bool inConditional = false;
            bool includeSection = true;
            bool skipUntilBlank = false; // 跳過非命盤主星的引文段落，直到空行

            foreach (var rawLine in lines)
            {
                var line = rawLine.TrimEnd();
                var trigger = KbDetectSectionTriggerTyped(line);
                if (trigger != null)
                {
                    skipUntilBlank = false; // 條件觸發行清除跳過狀態
                    var (triggers, isSameGong) = trigger.Value;
                    inConditional = true;
                    if (triggers.Count == 0)
                        includeSection = true;  // 通用格局，無特定星要求
                    else
                    {
                        // 同宮：星必須在同一宮位；會照/相夾：星在整個命盤中存在即可
                        var checkSet = isSameGong ? palaceStars : allChartStars;
                        includeSection = triggers.Any(s => checkSet.Contains(s));
                    }
                    // 若 trigger 標題行本身含有此宮不存在的副星 → 整個 section 不顯示
                    // 例：「文曲會照」中 '曲' 不在此宮 → 壓制整個段落，避免內容出現而標題被過濾
                    if (includeSection && KbContentLineHasAbsentSecondaryStar(line, palaceStars))
                        includeSection = false;
                }
                else if (!inConditional)
                {
                    // 在條件觸發區之前，若段落首行以不在命盤的主星開頭，跳過整個段落
                    if (string.IsNullOrWhiteSpace(line))
                        skipUntilBlank = false;
                    else if (!skipUntilBlank)
                    {
                        foreach (var star in KnownMainStarFullNames)
                        {
                            if (line.StartsWith(star) && !allChartStars.Contains(star) && !palaceStars.Contains(star))
                            {
                                skipUntilBlank = true;
                                break;
                            }
                        }
                    }
                }

                if (!inConditional && skipUntilBlank)
                    continue;

                if (!inConditional || includeSection)
                {
                    // 在條件段落內，過濾行內任意 [主星組合] 包含宮位外主星的行
                    // 例：丑未二垣[武曲貪狼]守命、[武貪]化忌 → 若貪狼不在宮位則跳過
                    if (inConditional && includeSection && KbLineHasUnavailableCombination(line, palaceStars))
                        continue;
                    // 過濾描述「副星同度/加臨」但該副星不在此宮的內容行
                    // 例：「巳亥[廉貪]有火鈴同度，...」→ 宮位無火星/鈴星時跳過
                    // 若非條件段落內（即獨立副星段落標題如「文曲」），同時設 skipUntilBlank
                    // 確保後續不含副星字的內容行也一起被略過
                    if (KbContentLineHasAbsentSecondaryStar(line, palaceStars))
                    {
                        if (!inConditional) skipUntilBlank = true;
                        continue;
                    }
                    result.AppendLine(line);
                }
            }
            return result.ToString().TrimEnd();
        }

        // --- KB Helper: 行內任意 [主星組合] 括號包含宮位外主星 → 應過濾 ---
        // 用於大限命宮 KB 內容：避免顯示武貪/武相等宮位外組合描述
        private static bool KbLineHasUnavailableCombination(string line, HashSet<string> palaceStars)
        {
            int pos = 0;
            while (pos < line.Length)
            {
                int s = line.IndexOf('[', pos);
                if (s < 0) break;
                int e = line.IndexOf(']', s);
                if (e < 0) break;
                string inner = line.Substring(s + 1, e - s - 1);
                // 只針對主星縮寫/全名，略過[遷移宮]等宮位名括號
                foreach (var (abbr, fullName) in StarAbbrToFull)
                {
                    if (!Array.Exists(KnownMainStarFullNames, n => n == fullName)) continue;
                    if (!inner.Contains(abbr) && !inner.Contains(fullName)) continue;
                    if (!palaceStars.Contains(fullName) && !palaceStars.Contains(abbr))
                        return true; // 括號內有宮位外主星
                }
                pos = e + 1;
            }
            return false;
        }

        // --- KB Helper: 判斷內容行是否提及此宮沒有的副星 ---
        // 逐字掃描句子，遇到 AllSecondaryStarSingleChars 中的字即視為副星引用
        // 只要有任一副星在此宮（palaceStars 含 JSON 原始單字）→ 保留；全部不在 → 過濾
        // 寧可過濾，不要顯示不該出現的內容
        private static bool KbContentLineHasAbsentSecondaryStar(string line, HashSet<string> palaceStars)
        {
            bool anyPresent = false;
            bool anyAbsent  = false;

            foreach (char c in line)
            {
                if (!AllSecondaryStarSingleChars.Contains(c)) continue;
                string s = c.ToString();
                if (palaceStars.Contains(s))
                    anyPresent = true;
                else
                    anyAbsent = true;
                if (anyPresent) break; // 已確認有副星在宮，提早結束
            }

            // 句中出現副星字，且全部不在此宮 → 過濾
            return anyAbsent && !anyPresent;
        }

        // --- KB Helper: 主星全名列表（用於觸發偵測）---
        private static readonly string[] KnownMainStarFullNames = new[]
        {
            "紫微","天機","太陽","武曲","天同","廉貞","天府","太陰","貪狼","巨門","天相","天梁","七殺","破軍"
        };

        // --- KB Helper: 主星縮寫→全名（用於四化組合展開）---
        private static readonly Dictionary<string, string> StarAbbrToFull = new()
        {
            {"紫","紫微"},{"機","天機"},{"陽","太陽"},{"武","武曲"},{"同","天同"},{"廉","廉貞"},
            {"府","天府"},{"陰","太陰"},{"貪","貪狼"},{"巨","巨門"},{"相","天相"},{"梁","天梁"},
            {"殺","七殺"},{"破","破軍"},{"昌","文昌"},{"曲","文曲"},{"輔","左輔"},{"弼","右弼"}
        };

        // --- KB Helper: 將先天四化組合（如「貪狼化忌」）加入 chartStars，供主星觸發條件過濾使用 ---
        private static void KbAddActiveTransformations(HashSet<string> chartStars, string yearStem)
        {
            if (!YearStemSiHuaMap.TryGetValue(yearStem, out var siHua)) return;
            string[] abbrList = { siHua.lu, siHua.quan, siHua.ke, siHua.ji };
            string[] huaList  = { "化祿", "化權", "化科", "化忌" };
            for (int i = 0; i < 4; i++)
            {
                string abbr = abbrList[i];
                string hua  = huaList[i];
                chartStars.Add(abbr + hua);  // 縮寫形式，如「貪化忌」
                if (StarAbbrToFull.TryGetValue(abbr, out var full))
                    chartStars.Add(full + hua);  // 全名形式，如「貪狼化忌」
            }
        }

        // --- KB Helper: 偵測行是否為「條件段落標題」，回傳 (觸發星列表, 是否為同宮類型) ---
        private static (List<string> triggers, bool isSameGong)? KbDetectSectionTriggerTyped(string line)
        {
            if (string.IsNullOrWhiteSpace(line) || line.Length > 35) return null;
            bool isSameGong = line.Contains("同宮");
            bool hasTrigger = isSameGong || line.Contains("會照") || line.Contains("相夾") ||
                              line.Contains("化忌") || line.Contains("化祿") || line.Contains("化科") ||
                              line.Contains("化權") || line.Contains("入廟") || line.Contains("獨坐");
            bool hasGe = line.Contains("格：") || (line.EndsWith("格") && line.Length < 18);
            if (!hasTrigger && !hasGe) return null;
            var found = new List<string>();
            // 先檢查副星
            foreach (var star in KnownSecondaryStarNames)
                if (line.Contains(star)) found.Add(star);
            // 再檢查主星
            if (found.Count == 0)
            {
                bool hasHua = line.Contains("化忌") || line.Contains("化祿") ||
                              line.Contains("化科") || line.Contains("化權");
                // 是否為「必須在同宮」型觸發（獨坐/坐命/守命/同宮 等）
                bool isPalaceBound = isSameGong || line.Contains("獨坐") ||
                                     line.Contains("坐命") || line.Contains("守命") ||
                                     line.Contains("守度") || line.Contains("入命") ||
                                     line.Contains("坐宮");
                foreach (var mainStar in KnownMainStarFullNames)
                {
                    if (!line.Contains(mainStar)) continue;
                    if (hasHua)
                    {
                        // 主星+四化：觸發 key = "主星名化X"，如「貪狼化權」；對照 allChartStars（含活躍四化）
                        string huaType = line.Contains("化忌") ? "化忌" :
                                         line.Contains("化祿") ? "化祿" :
                                         line.Contains("化科") ? "化科" : "化權";
                        found.Add(mainStar + huaType);
                    }
                    else if (isPalaceBound)
                    {
                        // 主星+獨坐/同宮等：必須在該宮才顯示；對照 palaceStars
                        found.Add(mainStar);
                    }
                    // 主星單純 會照/相夾 (不含化X) → 主星必在整個命盤中，一律顯示（不過濾）
                }
                // 檢查 [X] / [XY] 括號格式的主星段落標題，如「[廉貪]...」「[七殺]...」
                if (found.Count == 0 && line.StartsWith("["))
                {
                    int closeBracket = line.IndexOf(']');
                    if (closeBracket > 0 && closeBracket <= 8)
                    {
                        string inside = line.Substring(1, closeBracket - 1);
                        foreach (var kv in StarAbbrToFull)
                            if (inside.Contains(kv.Key)) found.Add(kv.Value);
                        // 若括號內只有全名（如[天相]）
                        if (found.Count == 0)
                            foreach (var mainStar in KnownMainStarFullNames)
                                if (inside.Contains(mainStar)) found.Add(mainStar);
                        // [X] 型一律視為同宮條件（必須在此宮才有效）
                        if (found.Count > 0) return (found, true);
                    }
                }
                // isSameGong 標記：若觸發來自 isPalaceBound → true；來自 化X → false（對照 allChartStars）
                if (found.Count > 0 && !hasHua) return (found, true);
            }
            if (hasGe && found.Count == 0)
            {
                foreach (var kv in FormationStarMap)
                    if (line.Contains(kv.Key)) { found.AddRange(kv.Value); break; }
                if (found.Count == 0) return (new List<string>(), false); // 通用格，無條件顯示
            }
            return found.Count > 0 ? (found, isSameGong) : null;
        }

        // 星曜縮寫 → 可能的全名列表（JSON 可能用全名或縮寫）
        private static readonly Dictionary<string, string[]> StarAbbrAliases = new()
        {
            {"輔", new[]{"輔","左輔","左"}}, {"弼", new[]{"弼","右弼","右"}},
            {"昌", new[]{"昌","文昌"}}, {"曲", new[]{"曲","文曲"}},
            {"魁", new[]{"魁","天魁"}}, {"鉞", new[]{"鉞","天鉞"}},
            {"馬", new[]{"馬","天馬"}}, {"祿", new[]{"祿","祿存"}},
            {"羊", new[]{"羊","擎羊"}}, {"陀", new[]{"陀","陀羅"}},
            {"火", new[]{"火","火星"}}, {"鈴", new[]{"鈴","鈴星"}},
            {"廉", new[]{"廉","廉貞"}}, {"破", new[]{"破","破軍"}},
            {"武", new[]{"武","武曲"}}, {"同", new[]{"同","天同"}},
            {"機", new[]{"機","天機"}}, {"陽", new[]{"陽","太陽"}},
            {"陰", new[]{"陰","太陰"}}, {"貪", new[]{"貪","貪狼"}},
            {"巨", new[]{"巨","巨門"}}, {"相", new[]{"相","天相"}},
            {"梁", new[]{"梁","天梁"}}, {"殺", new[]{"殺","七殺"}},
            {"府", new[]{"府","天府"}}, {"紫", new[]{"紫","紫微"}},
        };

        // --- KB Helper: 依星曜縮寫找宮位名（返回正規化 XXX宮 格式），支援全名/縮寫兩種 JSON 格式 ---
        private static string KbFindPalaceByStarAbbr(JsonElement palaces, string starAbbr)
        {
            if (palaces.ValueKind != JsonValueKind.Array || string.IsNullOrEmpty(starAbbr)) return "";
            var aliases = StarAbbrAliases.TryGetValue(starAbbr, out var al) ? al : new[] { starAbbr };
            foreach (var p in palaces.EnumerateArray())
            {
                string pname = p.TryGetProperty("palaceName", out var pn) ? pn.GetString() ?? "" :
                               p.TryGetProperty("name", out var n2) ? n2.GetString() ?? "" : "";
                foreach (var key in new[] { "majorStars", "mainStars", "secondaryStars" })
                {
                    if (p.TryGetProperty(key, out var stars) && stars.ValueKind == JsonValueKind.Array &&
                        stars.EnumerateArray().Any(s => aliases.Contains(s.GetString() ?? "")))
                        return KbNormalizePalaceName(pname);
                }
            }
            return "";
        }

        // --- KB Helper: 年干先天四化對照表（14主星+輔曜縮寫）---
        private static readonly Dictionary<string, (string lu, string quan, string ke, string ji)> YearStemSiHuaMap = new()
        {
            {"甲", ("廉","破","武","陽")}, {"乙", ("機","梁","紫","陰")},
            {"丙", ("同","機","昌","廉")}, {"丁", ("陰","同","機","巨")},
            {"戊", ("貪","陰","弼","機")}, {"己", ("武","貪","梁","曲")},
            {"庚", ("陽","武","陰","同")}, {"辛", ("巨","陽","曲","昌")},
            {"壬", ("梁","紫","輔","武")}, {"癸", ("破","巨","陰","貪")},
        };

        // --- KB Helper: 取先天四化落宮名 ---
        private static string KbGetSiHuaPalace(string yearStem, string siHuaType, JsonElement palaces)
        {
            if (!YearStemSiHuaMap.TryGetValue(yearStem, out var siHua)) return "";
            string starAbbr = siHuaType switch
            {
                "化祿" => siHua.lu, "化權" => siHua.quan,
                "化科" => siHua.ke, "化忌" => siHua.ji, _ => ""
            };
            return string.IsNullOrEmpty(starAbbr) ? "" : KbFindPalaceByStarAbbr(palaces, starAbbr);
        }

        // --- KB Helper: 先天四化 KB 查詢 ---
        private async Task<string> KbSiHuaQuery(string yearStem, string siHuaType, JsonElement palaces)
        {
            if (!YearStemSiHuaMap.TryGetValue(yearStem, out var siHua)) return "";
            string starAbbr = siHuaType switch
            {
                "化祿" => siHua.lu, "化權" => siHua.quan,
                "化科" => siHua.ke, "化忌" => siHua.ji, _ => ""
            };
            if (string.IsNullOrEmpty(starAbbr)) return "";
            string palaceName = KbFindPalaceByStarAbbr(palaces, starAbbr);
            if (string.IsNullOrEmpty(palaceName)) return "";
            string palaceShort = palaceName.TrimEnd('宮');
            string title = $"{siHuaType}星十二宮";
            return await KbQuery($"SELECT COALESCE(\"ResultText\",'') AS \"Value\" FROM \"FortuneRules\" WHERE \"SourceFile\"='先天四化.docx' AND \"Title\"='{title}' AND \"ConditionText\" LIKE '%{palaceShort}%' LIMIT 1");
        }

        // --- KB Helper: 取宮位宮干 ---
        private static string KbGetPalaceStem(JsonElement palaces, string palaceName)
        {
            if (palaces.ValueKind != JsonValueKind.Array) return "";
            foreach (var p in palaces.EnumerateArray())
            {
                string pname = p.TryGetProperty("palaceName", out var pn) ? pn.GetString() ?? "" :
                               p.TryGetProperty("name", out var n2) ? n2.GetString() ?? "" : "";
                if (!KbPalaceSame(pname, palaceName)) continue;
                return p.TryGetProperty("palaceStem", out var ps) ? ps.GetString() ?? "" : "";
            }
            return "";
        }

        // --- KB Helper: 宮位四化查詢（宮干→化星→落宮→查 FortuneRules）---
        // 回傳 (目標宮位名, 內容文字)
        private async Task<(string targetPalace, string content)> KbGongWeiSiHuaQuery(JsonElement palaces, string sourcePalaceName, string siHuaType)
        {
            string palaceStem = KbGetPalaceStem(palaces, sourcePalaceName);
            if (string.IsNullOrEmpty(palaceStem)) return ("", "");
            if (!YearStemSiHuaMap.TryGetValue(palaceStem, out var siHua)) return ("", "");
            string starAbbr = siHuaType switch
            {
                "化祿" => siHua.lu, "化權" => siHua.quan,
                "化科" => siHua.ke, "化忌" => siHua.ji, _ => ""
            };
            if (string.IsNullOrEmpty(starAbbr)) return ("", "");
            string targetPalace = KbFindPalaceByStarAbbr(palaces, starAbbr);
            if (string.IsNullOrEmpty(targetPalace)) return ("", "");
            // 宮位四化.docx 格式：Title='{源宮}四化飛星－－{化type}'
            // 自化（落本宮）格式：ResultText 以 "{源宮}自{化type}：" 開頭
            // 飛出（落他宮）格式：ResultText 以 "{源宮}{化type}入{目標短名}" 開頭
            string titleKey = $"{sourcePalaceName}四化飛星－－{siHuaType}";
            string resultPrefix = KbPalaceSame(targetPalace, sourcePalaceName)
                ? $"{sourcePalaceName}自{siHuaType}"
                : $"{sourcePalaceName}{siHuaType}入{targetPalace.TrimEnd('宮')}";
            string content = await KbQuery($"SELECT COALESCE(\"ResultText\",'') AS \"Value\" FROM \"FortuneRules\" WHERE \"SourceFile\"='宮位四化.docx' AND \"Title\"='{titleKey}' AND \"ResultText\" LIKE '{resultPrefix}%' LIMIT 1");
            return (targetPalace, content);
        }

        // --- KB Helper: 主星入宮查詢（6/7/8 主星文件）---
        // File 6(殺破貪) & 7(機同梁): ResultText 以 "{星名}星入{宮}宮：" 開頭
        // File 8(相廉武巨): Title = "{星名}星入{宮短名}"，ResultText = 描述
        private static readonly Dictionary<string, (string fullName, string file, bool useTitle)> StarInPalaceMap = new()
        {
            {"殺", ("七殺",  "6紫微斗数主星杀破狼入十二宮.docx", false)},
            {"破", ("破軍",  "6紫微斗数主星杀破狼入十二宮.docx", false)},
            {"貪", ("貪狼",  "6紫微斗数主星杀破狼入十二宮.docx", false)},
            {"機", ("天機",  "7紫微斗数主星机同梁入十二宮.docx",  false)},
            {"同", ("天同",  "7紫微斗数主星机同梁入十二宮.docx",  false)},
            {"梁", ("天梁",  "7紫微斗数主星机同梁入十二宮.docx",  false)},
            {"相", ("天相",  "8紫微斗数主星相廉武巨入十二宮.docx", true)},
            {"廉", ("廉貞",  "8紫微斗数主星相廉武巨入十二宮.docx", true)},
            {"武", ("武曲",  "8紫微斗数主星相廉武巨入十二宮.docx", true)},
            {"巨", ("巨門",  "8紫微斗数主星相廉武巨入十二宮.docx", true)},
        };

        // 5新.docx 格式：一條 ResultText 含所有十二宮，逐行提取
        private static readonly Dictionary<string, string> NewFormatStarMap = new()
        {
            {"府", "天府"},
            {"陽", "太陽"},
            {"陰", "太陰"},
        };

        // 雙星組合 Key 對照（兩星縮寫 → 雙星文件 Title Key）
        private static readonly Dictionary<string, string> DoubleStarKeyMap = new()
        {
            {"紫破","紫破"},{"破紫","紫破"}, {"紫府","紫府"},{"府紫","紫府"},
            {"紫相","紫相"},{"相紫","紫相"}, {"紫殺","紫殺"},{"殺紫","紫殺"},
            {"紫貪","紫貪"},{"貪紫","紫貪"},
            {"武府","武府"},{"府武","武府"}, {"武相","武相"},{"相武","武相"},
            {"武破","武破"},{"破武","武破"}, {"武殺","武殺"},{"殺武","武殺"},
            {"武貪","武貪"},{"貪武","武貪"},
            {"廉府","廉府"},{"府廉","廉府"}, {"廉相","廉相"},{"相廉","廉相"},
            {"廉破","廉破"},{"破廉","廉破"}, {"廉殺","廉殺"},{"殺廉","廉殺"},
            {"廉貪","廉貪"},{"貪廉","廉貪"},
            {"同巨","同巨"},{"巨同","同巨"}, {"同梁","同梁"},{"梁同","同梁"},
            {"同陰","同陰"},{"陰同","同陰"},
            {"機巨","機巨"},{"巨機","機巨"}, {"機梁","機梁"},{"梁機","機梁"},
            {"機陰","機陰"},{"陰機","機陰"},
            {"陽梁","陽梁"},{"梁陽","陽梁"},
            {"巨陽","巨日"},{"陽巨","巨日"},
            {"陽陰","日月"},{"陰陽","日月"},
        };

        private async Task<string> KbQueryStarInPalace(JsonElement palaces, string palaceName)
        {
            if (palaces.ValueKind != JsonValueKind.Array) return "";
            var stars = new List<string>();
            foreach (var p in palaces.EnumerateArray())
            {
                string pn = p.TryGetProperty("palaceName", out var pp) ? pp.GetString() ?? "" :
                            p.TryGetProperty("name", out var nn) ? nn.GetString() ?? "" : "";
                if (!KbPalaceSame(pn, palaceName)) continue;
                foreach (var key in new[] { "majorStars", "mainStars" })
                {
                    if (!p.TryGetProperty(key, out var s) || s.ValueKind != JsonValueKind.Array) continue;
                    foreach (var star in s.EnumerateArray())
                    {
                        var n = star.GetString() ?? "";
                        if (!string.IsNullOrEmpty(n)) stars.Add(n);
                    }
                    break;
                }
                break;
            }
            if (stars.Count == 0) return "";

            string palaceShort = KbNormalizePalaceName(palaceName).TrimEnd('宮');
            var parts = new StringBuilder();
            foreach (var star in stars)
            {
                // 5新.docx 格式：天府/太陽/太陰
                if (NewFormatStarMap.TryGetValue(star, out var fullStarName))
                {
                    string title = $"{fullStarName}星入十二宮";
                    string fullText = await KbQuery($"SELECT COALESCE(\"ResultText\",'') AS \"Value\" FROM \"FortuneRules\" WHERE \"SourceFile\"='5新 紫微斗数主星府陽陰四大主星入十二宮.docx' AND \"Title\"='{title}' LIMIT 1");
                    string line = KbExtractNewFormatStarLine(fullText, fullStarName, palaceShort);
                    if (!string.IsNullOrWhiteSpace(line)) parts.AppendLine($"【{fullStarName}】{line.Trim()}");
                    continue;
                }
                if (!StarInPalaceMap.TryGetValue(star, out var info)) continue;
                string result;
                if (info.useTitle)
                    result = await KbQuery($"SELECT COALESCE(\"ResultText\",'') AS \"Value\" FROM \"FortuneRules\" WHERE \"SourceFile\"='{info.file}' AND \"Title\" LIKE '{info.fullName}星入{palaceShort}%' LIMIT 1");
                else
                    result = await KbQuery($"SELECT COALESCE(\"ResultText\",'') AS \"Value\" FROM \"FortuneRules\" WHERE \"SourceFile\"='{info.file}' AND \"ResultText\" LIKE '{info.fullName}星入{palaceShort}宮：%' LIMIT 1");
                if (string.IsNullOrWhiteSpace(result)) continue;
                if (!info.useTitle && result.Contains('：'))
                    result = result.Substring(result.IndexOf('：') + 1);
                parts.AppendLine($"【{info.fullName}】{result.Trim()}");
            }
            return parts.ToString().Trim();
        }

        // 從 5新.docx 的多宮合併 ResultText 中提取特定宮位描述
        private static string KbExtractNewFormatStarLine(string fullText, string starFullName, string palaceShort)
        {
            string linePrefix = $"{starFullName}星入{palaceShort}宮：";
            foreach (var line in fullText.Split('\n'))
            {
                var trimmed = line.Trim();
                if (trimmed.StartsWith(linePrefix))
                    return trimmed.Substring(linePrefix.Length).Trim();
            }
            return "";
        }

        // 查詢雙星組合入宮（紫微斗数双星入十二宮(全).docx）
        private async Task<string> KbQueryDoubleStarInPalace(JsonElement palaces, string palaceName)
        {
            if (palaces.ValueKind != JsonValueKind.Array) return "";
            var stars = new List<string>();
            foreach (var p in palaces.EnumerateArray())
            {
                string pn = p.TryGetProperty("palaceName", out var pp) ? pp.GetString() ?? "" :
                            p.TryGetProperty("name", out var nn) ? nn.GetString() ?? "" : "";
                if (!KbPalaceSame(pn, palaceName)) continue;
                foreach (var key in new[] { "majorStars", "mainStars" })
                {
                    if (!p.TryGetProperty(key, out var s) || s.ValueKind != JsonValueKind.Array) continue;
                    foreach (var star in s.EnumerateArray())
                    {
                        var n = star.GetString() ?? "";
                        if (!string.IsNullOrEmpty(n)) stars.Add(n);
                    }
                    break;
                }
                break;
            }
            if (stars.Count < 2) return "";

            string doubleKey = "";
            for (int i = 0; i < stars.Count && string.IsNullOrEmpty(doubleKey); i++)
                for (int j = i + 1; j < stars.Count && string.IsNullOrEmpty(doubleKey); j++)
                    if (DoubleStarKeyMap.TryGetValue(stars[i] + stars[j], out var k)) doubleKey = k;

            if (string.IsNullOrEmpty(doubleKey)) return "";

            // 官祿 → 事業（雙星文件使用「事業」）
            string palaceNorm = KbNormalizePalaceName(palaceName).Replace("官祿", "事業");
            string titlePalace = palaceNorm switch
            {
                "夫妻宮" or "子女宮" or "父母宮" or "兄弟宮" or "交友宮" => "六親宮位",
                _ => palaceNorm
            };

            string result = await KbQuery($"SELECT COALESCE(\"ResultText\",'') AS \"Value\" FROM \"FortuneRules\" WHERE \"SourceFile\"='紫微斗数双星入十二宮(全).docx' AND \"Title\" IN ('{doubleKey}居{titlePalace}', '{doubleKey}{titlePalace}') LIMIT 1");
            return result?.Trim() ?? "";
        }

        // 查詢輔星（吉星/煞星）入宮
        private async Task<string> KbQueryMinorStarsInPalace(JsonElement palaces, string palaceName)
        {
            if (palaces.ValueKind != JsonValueKind.Array) return "";
            var secondarySet = new HashSet<string>();
            var goodSet = new HashSet<string>();
            var badSet = new HashSet<string>();

            foreach (var p in palaces.EnumerateArray())
            {
                string pn = p.TryGetProperty("palaceName", out var pp) ? pp.GetString() ?? "" :
                            p.TryGetProperty("name", out var nn) ? nn.GetString() ?? "" : "";
                if (!KbPalaceSame(pn, palaceName)) continue;
                void FillSet(HashSet<string> set, string jsonKey)
                {
                    if (!p.TryGetProperty(jsonKey, out var arr) || arr.ValueKind != JsonValueKind.Array) return;
                    foreach (var s in arr.EnumerateArray()) { var n = s.GetString() ?? ""; if (!string.IsNullOrEmpty(n)) set.Add(n); }
                }
                FillSet(secondarySet, "secondaryStars");
                FillSet(goodSet, "goodStars");
                FillSet(badSet, "badStars");
                break;
            }

            // 官祿 → 事業（副星/煞星文件使用「事業」）
            string palaceNormForMinor = KbNormalizePalaceName(palaceName).Replace("官祿", "事業");
            string palaceShortNoGong = palaceNormForMinor.TrimEnd('宮');
            string palaceFull = palaceNormForMinor;
            var sb = new StringBuilder();

            bool HasStar(HashSet<string> set, string abbr) => set.Any(s => s == abbr || s.StartsWith(abbr + "("));

            // 吉星（9.docx）：title 不含宮後綴
            var luckyGroups = new (string[] abbrs, string titleKey)[]
            {
                (new[]{ "昌","曲" }, "文昌星、文曲星"),
                (new[]{ "左","右" }, "左輔星、右弼星"),
                (new[]{ "魁","鉞" }, "天魁星、天鉞星"),
                (new[]{ "祿","馬" }, "䘵存星、天馬星"),
            };
            foreach (var (abbrs, titleKey) in luckyGroups)
            {
                if (!abbrs.Any(a => HasStar(secondarySet, a) || HasStar(goodSet, a))) continue;
                string r = await KbQuery($"SELECT COALESCE(\"ResultText\",'') AS \"Value\" FROM \"FortuneRules\" WHERE \"SourceFile\"='9紫微斗数辅星-吉星入十二宮.docx' AND \"Title\"='{titleKey}入{palaceShortNoGong}' LIMIT 1");
                if (!string.IsNullOrWhiteSpace(r)) sb.AppendLine(r.Trim());
            }

            // 煞星（4.docx）：title 含宮後綴
            var evilGroups = new (string abbr, HashSet<string> set, string starName)[]
            {
                ("羊", badSet, "擎羊星"),
                ("陀", badSet, "陀羅星"),
                ("火", badSet, "火星"),
                ("鈴", badSet, "鈴星"),
                ("劫", secondarySet, "地劫星"),
                ("空", secondarySet, "地空星"),
            };
            foreach (var (abbr, set, starName) in evilGroups)
            {
                if (!HasStar(set, abbr)) continue;
                string r = await KbQuery($"SELECT COALESCE(\"ResultText\",'') AS \"Value\" FROM \"FortuneRules\" WHERE \"SourceFile\"='4紫微斗数煞星入十二宮.docx' AND \"Title\"='{starName}入{palaceFull}' LIMIT 1");
                if (!string.IsNullOrWhiteSpace(r)) sb.AppendLine(r.Trim());
            }

            return sb.ToString().Trim();
        }

        // --- KB Helper: 取地支主氣十神 ---
        private static string KbGetFirstHiddenSS(JsonElement pillar)
        {
            if (!pillar.TryGetProperty("hiddenStemLiuShen", out var arr) || arr.ValueKind != JsonValueKind.Array) return "";
            return arr.EnumerateArray().FirstOrDefault().GetString() ?? "";
        }

        // --- KB Helper: 取當前八字大運 ---
        private static (string stem, string branch, string liuShen, int start, int end) KbGetCurrentLuck(JsonElement root, int currentAge)
        {
            if (!root.TryGetProperty("baziLuckCycles", out var cycles) || cycles.ValueKind != JsonValueKind.Array)
                return ("", "", "", 0, 0);
            foreach (var c in cycles.EnumerateArray())
            {
                int s = c.TryGetProperty("startAge", out var sa) ? sa.GetInt32() : 0;
                int e = c.TryGetProperty("endAge",   out var ea) ? ea.GetInt32() : 0;
                if (currentAge >= s && currentAge <= e)
                {
                    string stem = c.TryGetProperty("heavenlyStem",  out var hs) ? hs.GetString() ?? "" : "";
                    string bran = c.TryGetProperty("earthlyBranch", out var eb) ? eb.GetString() ?? "" : "";
                    string ls   = c.TryGetProperty("liuShen",       out var l)  ? l.GetString()  ?? "" : "";
                    return (stem, bran, ls, s, e);
                }
            }
            return ("", "", "", 0, 0);
        }

        // --- KB Helper: 取紫微大限宮位 ---
        private static string KbGetLuckPalace(JsonElement palaces, int currentAge)
        {
            if (palaces.ValueKind != JsonValueKind.Array) return "";
            foreach (var p in palaces.EnumerateArray())
            {
                string range = p.TryGetProperty("decadeAgeRange", out var dr) ? dr.GetString() ?? "" : "";
                var parts = range.Split('-');
                if (parts.Length == 2 && int.TryParse(parts[0], out int ps) && int.TryParse(parts[1], out int pe)
                    && currentAge >= ps && currentAge <= pe)
                {
                    string pn2 = p.TryGetProperty("palaceName", out var pn) ? pn.GetString() ?? "" : "";
                    return KbNormalizePalaceName(pn2);
                }
            }
            return "";
        }

        // --- KB Helper: 附加六神四柱數 ---
        private static void KbAppendGd(StringBuilder sb, string label, string content)
        {
            if (!string.IsNullOrWhiteSpace(content))
                sb.AppendLine($"【{label}】{content}");
        }

        // --- KB Helper: 剝除 HTML 標籤 ---
        private static string KbStripHtml(string input)
            => Regex.Replace(input, "<.*?>", " ").Replace("  ", " ").Trim();

        // --- KB Helper: 日主天干→五行 ---
        private static string KbStemToElement(string stem) => stem switch
        {
            "甲" or "乙" => "木",
            "丙" or "丁" => "火",
            "戊" or "己" => "土",
            "庚" or "辛" => "金",
            "壬" or "癸" => "水",
            _ => ""
        };

        // --- KB Helper: 月支→季節 ---
        private static string KbBranchToSeason(string branch) => branch switch
        {
            "寅" or "卯" or "辰" => "春",
            "巳" or "午" or "未" => "夏",
            "申" or "酉" or "戌" => "秋",
            "亥" or "子" or "丑" => "冬",
            _ => "春"
        };

        // --- KB Helper: 時支→astro_twoheader 欄位 ---
        private static string KbBranchToHourCol(string branch) => branch switch
        {
            "子" or "丑" => "D",
            "寅" or "卯" => "E",
            "辰" or "巳" => "G",
            "午" or "未" => "H",
            "申" or "酉" => "J",
            "戌" or "亥" => "K",
            _ => "D"
        };

        // --- KB Helper: 農曆轉中文日期（1963415 → 一九六三年四月十五日）---
        private static string KbLunarToChineseStr(string lunar)
        {
            if (string.IsNullOrEmpty(lunar) || lunar.Length < 5) return lunar;
            string[] digits = { "零","一","二","三","四","五","六","七","八","九" };
            string year = lunar[..4];
            string rest = lunar[4..];
            if (!int.TryParse(rest, out _)) return lunar;

            int month, day;
            if (rest.Length == 1) { month = int.Parse(rest); day = 0; }
            else if (rest.Length == 2) { month = int.Parse(rest[..1]); day = int.Parse(rest[1..]); }
            else if (rest.Length == 3)
            {
                // Try 2-digit month first (10,11,12), else 1-digit
                if (int.TryParse(rest[..2], out int mm) && mm >= 10 && mm <= 12)
                { month = mm; day = int.Parse(rest[2..]); }
                else { month = int.Parse(rest[..1]); day = int.Parse(rest[1..]); }
            }
            else { month = int.Parse(rest[..2]); day = int.Parse(rest[2..]); }

            string yearCh = string.Concat(year.Select(c => digits[c - '0']));
            string[] monthNames = { "","正","二","三","四","五","六","七","八","九","十","十一","十二" };
            string[] dayNames   = { "","初一","初二","初三","初四","初五","初六","初七","初八","初九","初十",
                                     "十一","十二","十三","十四","十五","十六","十七","十八","十九","二十",
                                     "二十一","二十二","二十三","二十四","二十五","二十六","二十七","二十八","二十九","三十" };
            string monthCh = (month >= 1 && month <= 12) ? monthNames[month] : month.ToString();
            string dayCh   = (day >= 1 && day <= 30)     ? dayNames[day]     : day.ToString();
            return $"{yearCh}年{monthCh}月{dayCh}日";
        }

        // --- KB Helper: 陽曆日期格式化（去掉時間部分）---
        private static string KbFormatSolarDate(string solar)
        {
            if (string.IsNullOrEmpty(solar)) return "";
            // "1963-05-08T20:00:00" → "1963年05月08日"
            if (DateTime.TryParse(solar, out var dt))
                return $"{dt.Year}年{dt.Month:D2}月{dt.Day:D2}日";
            return solar.Split('T')[0];
        }

        // --- KB Helper: 解析四柱斷語 ---
        private static (string header, string nian, string yue, string ri, string shi) KbSplitPillars(string text)
        {
            if (string.IsNullOrEmpty(text)) return ("", "", "", "", "");
            var markers = new[] { "[年柱(根)]", "[月柱(苗)]", "[日柱(花)]", "[時柱(果)]" };
            int[] pos = markers.Select(m => text.IndexOf(m, StringComparison.Ordinal)).ToArray();
            string header = pos[0] > 0 ? text[..pos[0]].TrimEnd('-', ' ', '\n', '\r') : "";
            var parts = new string[4];
            for (int i = 0; i < 4; i++)
            {
                if (pos[i] < 0) continue;
                int start = pos[i] + markers[i].Length;
                int end = (i < 3 && pos[i + 1] > 0) ? pos[i + 1] : text.Length;
                parts[i] = text[start..end].TrimStart('：', ':', ' ', '\n', '\r').Trim();
            }
            return (header, parts[0], parts[1], parts[2], parts[3]);
        }

        // --- KB Helper: 格式化單一柱行 ---
        private static string KbFormatPillarLine(string content, string label, string context, bool isTimePillar)
        {
            if (string.IsNullOrEmpty(content)) return "";
            // 非時柱移除「生時XX，...。」錯誤文句
            if (!isTimePillar)
                content = Regex.Replace(content, @"生時[^，,]{1,6}[，,][^。]*。", "", RegexOptions.Singleline);
            content = KbStripHtml(content).Replace("  ", " ").Trim();
            return $"[{label}]({context})：{content}";
        }

        // --- KB Helper: 地支→生肖 ---
        // 移除 nianSiHuaXing 開頭的「X年生人」前綴（如「癸年生人」）
        private static string StripNianShengRen(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            return System.Text.RegularExpressions.Regex.Replace(s, @"^[\u4e00-\u9fa5]年生人", "").TrimStart();
        }

        private static string KbBranchToZodiac(string branch) => branch switch
        {
            "子" => "鼠", "丑" => "牛", "寅" => "虎", "卯" => "兔",
            "辰" => "龍", "巳" => "蛇", "午" => "馬", "未" => "羊",
            "申" => "猴", "酉" => "雞", "戌" => "狗", "亥" => "豬",
            _ => ""
        };

        // --- KB Helper: 十神縮寫展開（大運用）---
        private static string KbExpandLiuShen(string abbr) => abbr switch
        {
            "比" => "比肩", "劫" => "劫財", "食" => "食神", "傷" => "傷官",
            "財" => "偏財", "才" => "正財", "殺" => "七殺", "官" => "正官",
            "梟" => "偏印", "印" => "正印",
            _ => abbr
        };

        // ============================================================
        private async Task<string> CallGeminiApi(string prompt)
        {
            var apiKey = _config["Gemini:ApiKey"];
            var url = $"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={apiKey}";
            var payload = new { contents = new[] { new { parts = new[] { new { text = prompt } } } } };
            var response = await _httpClient.PostAsJsonAsync(url, payload);
            var rawJson = await response.Content.ReadAsStringAsync();
            _logger.LogInformation("Gemini HTTP {StatusCode}: {Preview}", (int)response.StatusCode, rawJson.Length > 300 ? rawJson[..300] : rawJson);
            var json = JsonSerializer.Deserialize<JsonElement>(rawJson);
            if (!json.TryGetProperty("candidates", out var candidates))
                throw new Exception($"Gemini 回傳錯誤: {rawJson[..Math.Min(300, rawJson.Length)]}");
            return candidates[0].GetProperty("content").GetProperty("parts")[0].GetProperty("text").GetString();
        }

        // ==========================================
        // ANALYZE-LIFELONG: 終身命書 (12 章，純規則引擎)
        // ==========================================

        [HttpGet("analyze-lifelong")]
        [Authorize]
        public async Task<IActionResult> GetLifelongAnalysis()
        {
            var identity = User.FindFirstValue(ClaimTypes.Email)
                         ?? User.FindFirstValue(ClaimTypes.Name)
                         ?? User.FindFirst("unique_name")?.Value;
            if (string.IsNullOrEmpty(identity))
                return Unauthorized(new { error = "請重新登入" });

            var user = await _context.Users.FirstOrDefaultAsync(u => u.UserName == identity || u.Email == identity);
            if (user == null) return BadRequest(new { error = "找不到用戶" });

            var userChart = await _context.UserCharts.FirstOrDefaultAsync(c => c.UserId == user.Id);
            if (userChart == null || string.IsNullOrEmpty(userChart.ChartJson))
                return BadRequest(new { error = "no_chart" });

            bool lfIsAdmin = string.Equals(user.Email, _config["Admin:Email"], StringComparison.OrdinalIgnoreCase);
            int lfSubId = -1;
            if (!lfIsAdmin)
            {
                var (lfOk, lfErr, lfSubIdVal) = await CheckSubscriptionQuota(user.Id, "BOOK_BAZI");
                if (!lfOk) return BadRequest(new { error = lfErr });
                lfSubId = lfSubIdVal;
            }

            try
            {
                var root = JsonDocument.Parse(userChart.ChartJson).RootElement;
                if (!root.TryGetProperty("bazi", out var bazi) && !root.TryGetProperty("baziInfo", out bazi))
                    return BadRequest(new { error = "命盤資料格式錯誤" });

                var yearP  = LfGetPillar(bazi, "yearPillar");
                var monthP = LfGetPillar(bazi, "monthPillar");
                var dayP   = LfGetPillar(bazi, "dayPillar");
                var timeP  = LfGetPillar(bazi, "timePillar");

                string yStem = LfPillarStem(yearP);   string yBranch = LfPillarBranch(yearP);
                string mStem = LfPillarStem(monthP);  string mBranch = LfPillarBranch(monthP);
                string dStem = LfPillarStem(dayP);    string dBranch = LfPillarBranch(dayP);
                string hStem = LfPillarStem(timeP);   string hBranch = LfPillarBranch(timeP);

                string yStemSS   = LfPillarStemSS(yearP);
                string mStemSS   = LfPillarStemSS(monthP);
                string hStemSS   = LfPillarStemSS(timeP);
                string yBranchSS = LfPillarBranchMainSS(yearP);
                string mBranchSS = LfPillarBranchMainSS(monthP);
                string dBranchSS = LfPillarBranchMainSS(dayP);
                string hBranchSS = LfPillarBranchMainSS(timeP);

                int birthYear = user.BirthYear ?? (DateTime.Today.Year - 30);
                int gender    = user.BirthGender ?? 1;
                var luckCycles = LfExtractLuckCycles(root);

                string dmElem  = KbStemToElement(dStem);
                var branches   = new[] { yBranch, mBranch, dBranch, hBranch };
                var wuXing     = LfCalcWuXingMatrix(yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch);
                double bodyPct = LfGetBodyStrengthPct(dmElem, wuXing);
                string bodyLabel = LfGetBodyStrengthLabel(bodyPct);
                string season  = LfGetSeason(mBranch);
                string seaLabel = LfGetSeasonLabel(mBranch);

                var (pattern, yongShenElem, fuYiElem, yongReason, tiaoHouElem) = LfDetectGeJuAndYongShen(
                    yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch,
                    dmElem, wuXing, bodyPct, season);
                string jiShenElem = LfGetJiShenElem(yongShenElem, dmElem, bodyPct, pattern);

                var chartStems = new[] { yStem, mStem, dStem, hStem };
                var scored = luckCycles.Select(lc =>
                {
                    int sc = LfCalcLuckScore(lc.stem, lc.branch, pattern, yongShenElem, fuYiElem, jiShenElem,
                        dmElem, bodyPct > 50, tiaoHouElem, season, branches, chartStems, dStem);
                    return (lc.stem, lc.branch, lc.liuShen, lc.startAge, lc.endAge, score: sc, level: LfLuckLevel(sc));
                }).ToList();

                string report = LfBuildReport(
                    yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch,
                    yStemSS, mStemSS, hStemSS, yBranchSS, mBranchSS, dBranchSS, hBranchSS,
                    dmElem, wuXing, bodyPct, bodyLabel, season, seaLabel,
                    pattern, yongShenElem, fuYiElem, yongReason, jiShenElem,
                    scored, gender, birthYear);

                var cycleData = scored.Select(c => new {
                    stem = c.stem, branch = c.branch, liuShen = c.liuShen,
                    startAge = c.startAge, endAge = c.endAge,
                    score = c.score, level = c.level
                }).ToList();

                var baziTable = new {
                    pillars = new[] {
                        new { label = "年", stem = yStem, branch = yBranch, stemSS = yStemSS,
                              naYin = LfPillarNaYin(yearP),
                              hiddenPairs = LfPillarHiddenPairs(yearP) },
                        new { label = "月", stem = mStem, branch = mBranch, stemSS = mStemSS,
                              naYin = LfPillarNaYin(monthP),
                              hiddenPairs = LfPillarHiddenPairs(monthP) },
                        new { label = "日", stem = dStem, branch = dBranch, stemSS = "元神",
                              naYin = LfPillarNaYin(dayP),
                              hiddenPairs = LfPillarHiddenPairs(dayP) },
                        new { label = "時", stem = hStem, branch = hBranch, stemSS = hStemSS,
                              naYin = LfPillarNaYin(timeP),
                              hiddenPairs = LfPillarHiddenPairs(timeP) },
                    }
                };

                // 建立天干地支喜忌結構化資料（供前端渲染）
                string tuneElemForTable = season == "冬" ? "火" : season == "夏" ? "水" : "";
                string jiYongElemForTable = LfElemOvercomeBy.GetValueOrDefault(yongShenElem, "");
                string ClsForTable(string elem) {
                    if (elem == jiShenElem) return "X";
                    if (elem == yongShenElem || elem == fuYiElem ||
                        (!string.IsNullOrEmpty(tuneElemForTable) && elem == tuneElemForTable)) return "○";
                    if (elem == jiYongElemForTable && jiYongElemForTable != jiShenElem) return "△忌";
                    return "△";
                }
                string[] allStems = { "甲","乙","丙","丁","戊","己","庚","辛","壬","癸" };
                string[] allBrs   = { "子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥" };
                var yongJiTable = new {
                    stems = allStems.Select(s => new {
                        stem = s,
                        elem = KbStemToElement(s),
                        shiShen = LfStemShiShen(s, dStem),
                        cls = ClsForTable(KbStemToElement(s))
                    }).ToArray(),
                    branches = allBrs.Select(br => {
                        string brElem = LfBranchHiddenRatio.TryGetValue(br, out var bh) && bh.Count > 0
                            ? KbStemToElement(bh[0].stem) : "-";
                        string brMs = LfBranchHiddenRatio.TryGetValue(br, out var bh2) && bh2.Count > 0 ? bh2[0].stem : "";
                        string brSS = !string.IsNullOrEmpty(brMs) ? LfStemShiShen(brMs, dStem) : "-";
                        return new {
                            branch = br,
                            elem = brElem,
                            shiShen = brSS,
                            cls = brElem != "-" ? ClsForTable(brElem) : "-",
                            inChart = branches.Contains(br)
                        };
                    }).ToArray()
                };

                // 九星氣學加成（純 KB）
                string lfNsSection = await NsBuildBirthSection(
                    user.BirthYear ?? DateTime.Today.Year - 30,
                    user.BirthMonth ?? 1,
                    user.BirthDay ?? 1,
                    user.BirthHour ?? 0,
                    user.BirthGender ?? 1);
                if (!string.IsNullOrEmpty(lfNsSection)) report += lfNsSection;

                if (!lfIsAdmin) await RecordSubscriptionClaim(user.Id, lfSubId, "BOOK_BAZI");
                await SaveUserReportAsync(user.Id, "lifelong", "終身命書", report,
                    new { birthYear = user.BirthYear, birthMonth = user.BirthMonth, birthDay = user.BirthDay, gender = user.BirthGender });
                return Ok(new { result = report, luckCycles = cycleData, baziTable, yongJiTable });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "終身命書失敗 User={User}", identity);
                return StatusCode(500, new { error = "終身命書生成失敗，請稍後再試", details = ex.Message });
            }
        }

        // ============================================================
        // ANALYZE-BAZI-ZIWEI: 玉洞子八字紫微命書（銅會員版）
        // 完整讀取 DB JSON（八字 + 紫微 palaces），八字主體 + 紫微補充
        // ============================================================

        [HttpGet("analyze-bazi-ziwei")]
        [Authorize]
        public async Task<IActionResult> GetBaziZiweiAnalysis()
        {
            var identity = User.FindFirstValue(ClaimTypes.Email)
                         ?? User.FindFirstValue(ClaimTypes.Name)
                         ?? User.FindFirst("unique_name")?.Value;
            if (string.IsNullOrEmpty(identity))
                return Unauthorized(new { error = "請重新登入" });

            var user = await _context.Users.FirstOrDefaultAsync(u => u.UserName == identity || u.Email == identity);
            if (user == null) return BadRequest(new { error = "找不到用戶" });

            bool bzIsAdmin = string.Equals(user.Email, _config["Admin:Email"], StringComparison.OrdinalIgnoreCase);
            int bzSubId = -1;
            if (!bzIsAdmin)
            {
                var (bzOk, bzErr, bzSubIdVal) = await CheckSubscriptionQuota(user.Id, "BOOK_BAZI");
                if (!bzOk) return BadRequest(new { error = bzErr });
                bzSubId = bzSubIdVal;
            }

            var userChart = await _context.UserCharts.FirstOrDefaultAsync(c => c.UserId == user.Id);
            if (userChart == null || string.IsNullOrEmpty(userChart.ChartJson))
                return BadRequest(new { error = "no_chart" });

            try
            {
                // === 完整讀取 DB JSON（包含 palaces）===
                var root = JsonDocument.Parse(userChart.ChartJson).RootElement;
                if (!root.TryGetProperty("bazi", out var bazi) && !root.TryGetProperty("baziInfo", out bazi))
                    return BadRequest(new { error = "命盤資料格式錯誤" });

                var yearP  = LfGetPillar(bazi, "yearPillar");
                var monthP = LfGetPillar(bazi, "monthPillar");
                var dayP   = LfGetPillar(bazi, "dayPillar");
                var timeP  = LfGetPillar(bazi, "timePillar");

                string yStem = LfPillarStem(yearP);   string yBranch = LfPillarBranch(yearP);
                string mStem = LfPillarStem(monthP);  string mBranch = LfPillarBranch(monthP);
                string dStem = LfPillarStem(dayP);    string dBranch = LfPillarBranch(dayP);
                string hStem = LfPillarStem(timeP);   string hBranch = LfPillarBranch(timeP);

                string yStemSS   = LfPillarStemSS(yearP);
                string mStemSS   = LfPillarStemSS(monthP);
                string hStemSS   = LfPillarStemSS(timeP);
                string yBranchSS = LfPillarBranchMainSS(yearP);
                string mBranchSS = LfPillarBranchMainSS(monthP);
                string dBranchSS = LfPillarBranchMainSS(dayP);
                string hBranchSS = LfPillarBranchMainSS(timeP);

                int birthYear = user.BirthYear ?? (DateTime.Today.Year - 30);
                int gender    = user.BirthGender ?? 1;
                var luckCycles = LfExtractLuckCycles(root);

                string dmElem  = KbStemToElement(dStem);
                var branches   = new[] { yBranch, mBranch, dBranch, hBranch };
                var wuXing     = LfCalcWuXingMatrix(yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch);
                double bodyPct = LfGetBodyStrengthPct(dmElem, wuXing);
                string bodyLabel = LfGetBodyStrengthLabel(bodyPct);
                string season  = LfGetSeason(mBranch);
                string seaLabel = LfGetSeasonLabel(mBranch);

                var (pattern, yongShenElem, fuYiElem, yongReason, tiaoHouElem) = LfDetectGeJuAndYongShen(
                    yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch,
                    dmElem, wuXing, bodyPct, season);
                string jiShenElem = LfGetJiShenElem(yongShenElem, dmElem, bodyPct, pattern);

                var chartStems = new[] { yStem, mStem, dStem, hStem };
                var scored = luckCycles.Select(lc =>
                {
                    int sc = LfCalcLuckScore(lc.stem, lc.branch, pattern, yongShenElem, fuYiElem, jiShenElem,
                        dmElem, bodyPct > 50, tiaoHouElem, season, branches, chartStems, dStem);
                    return (lc.stem, lc.branch, lc.liuShen, lc.startAge, lc.endAge, score: sc, level: LfLuckLevel(sc));
                }).ToList();

                // === 八字主體 ===
                string report = LfBuildReport(
                    yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch,
                    yStemSS, mStemSS, hStemSS, yBranchSS, mBranchSS, dBranchSS, hBranchSS,
                    dmElem, wuXing, bodyPct, bodyLabel, season, seaLabel,
                    pattern, yongShenElem, fuYiElem, yongReason, jiShenElem,
                    scored, gender, birthYear);

                // === 紫微斗數補充（從完整 JSON 讀取 palaces）===
                bool bzHasZiwei = root.TryGetProperty("palaces", out var bzPalaces)
                    && bzPalaces.ValueKind == JsonValueKind.Array;

                // Fallback：若存檔命盤無 palaces，從生辰重算
                if (!bzHasZiwei && user.BirthYear.HasValue && user.BirthMonth.HasValue && user.BirthDay.HasValue)
                {
                    try
                    {
                        var bzReq = new AstrologyRequest(
                            user.BirthYear.Value, user.BirthMonth.Value, user.BirthDay.Value,
                            user.BirthHour ?? 0, user.BirthMinute ?? 0,
                            user.BirthGender ?? 1, user.ChartName ?? user.UserName ?? "",
                            user.DateType ?? "solar");
                        var bzFresh = await _astrologyService.CalculateChartAsync(bzReq);
                        var bzOpts = new System.Text.Json.JsonSerializerOptions { PropertyNamingPolicy = System.Text.Json.JsonNamingPolicy.CamelCase };
                        string bzFreshJson = System.Text.Json.JsonSerializer.Serialize(bzFresh, bzOpts);
                        var bzFreshRoot = System.Text.Json.JsonDocument.Parse(bzFreshJson).RootElement;
                        bzHasZiwei = bzFreshRoot.TryGetProperty("palaces", out bzPalaces)
                            && bzPalaces.ValueKind == System.Text.Json.JsonValueKind.Array;
                        if (bzHasZiwei)
                        {
                            userChart.ChartJson = bzFreshJson;
                            userChart.UpdatedAt = DateTime.UtcNow;
                            await _context.SaveChangesAsync();
                        }
                    }
                    catch (Exception exBz) { _logger.LogWarning(exBz, "BaziZiwei Ziwei recalc fallback failed"); }
                }

                if (bzHasZiwei)
                {
                    string bzPos = KbGetZiweiPosition(bzPalaces);
                    var bzChartStars = KbGetAllChartStars(bzPalaces);
                    string bzFullContent = await KbZiweiFullQuery(bzPalaces, bzPos);
                    string bzMingGongStars = userChart.MingGongMainStars ?? "";
                    string bzMingZhu = root.TryGetProperty("mingZhu", out var mzBz) ? mzBz.GetString() ?? "" : "";
                    string bzShenZhu = root.TryGetProperty("shenZhu", out var szBz) ? szBz.GetString() ?? "" : "";
                    string bzWuXingJu = root.TryGetProperty("wuXingJuText", out var wxBz) ? wxBz.GetString() ?? "" : "";

                    // 加入先天四化至 chartStars 確保過濾正確
                    KbAddActiveTransformations(bzChartStars, yStem);

                    // 先天四化（年干）
                    string bzSiHuaLuPalace   = KbGetSiHuaPalace(yStem, "化祿", bzPalaces);
                    string bzSiHuaQuanPalace = KbGetSiHuaPalace(yStem, "化權", bzPalaces);
                    string bzSiHuaKePalace   = KbGetSiHuaPalace(yStem, "化科", bzPalaces);
                    string bzSiHuaJiPalace   = KbGetSiHuaPalace(yStem, "化忌", bzPalaces);
                    string bzSiHuaLu   = await KbSiHuaQuery(yStem, "化祿", bzPalaces);
                    string bzSiHuaQuan = await KbSiHuaQuery(yStem, "化權", bzPalaces);
                    string bzSiHuaKe   = await KbSiHuaQuery(yStem, "化科", bzPalaces);
                    string bzSiHuaJi   = await KbSiHuaQuery(yStem, "化忌", bzPalaces);

                    // 4.1 格局說明（命宮格局偵測）
                    string bzGeJu = "";
                    string bzMingBr = KbGetPalaceBranch(bzPalaces, "命宮");
                    var bzGjList = LfDetectZiweiGeJu(bzMingGongStars, bzMingBr, bzChartStars,
                        bzSiHuaLuPalace, bzSiHuaQuanPalace, bzSiHuaKePalace, bzPalaces);
                    if (bzGjList.Count > 0)
                    {
                        var bzGjSb = new System.Text.StringBuilder();
                        foreach (var gj in bzGjList)
                        {
                            string desc = await KbQuery($"SELECT COALESCE(\"ResultText\",'') AS \"Value\" FROM \"FortuneRules\" WHERE \"SourceFile\"='紫微格局說明.docx' AND \"Title\"='{gj}' LIMIT 1");
                            if (!string.IsNullOrEmpty(desc)) { bzGjSb.AppendLine($"【{gj}】"); bzGjSb.AppendLine(desc); bzGjSb.AppendLine(); }
                        }
                        bzGeJu = bzGjSb.ToString();
                    }

                    // 主星宮位內容（ziwei_patterns_144 過濾版）
                    string bzMing = KbFilterZiweiContent(KbExtractPalaceSection(bzFullContent, "命宮"), KbGetPalaceStarsSet(bzPalaces, "命宮"), bzChartStars);
                    string bzOff  = KbFilterZiweiContent(KbExtractPalaceSection(bzFullContent, "事業宮"), KbGetPalaceStarsSet(bzPalaces, "官祿"), bzChartStars);
                    string bzWlt  = KbFilterZiweiContent(KbExtractPalaceSection(bzFullContent, "財帛宮"), KbGetPalaceStarsSet(bzPalaces, "財帛"), bzChartStars);
                    string bzSps  = KbFilterZiweiContent(KbExtractPalaceSection(bzFullContent, "夫妻宮"), KbGetPalaceStarsSet(bzPalaces, "夫妻"), bzChartStars);
                    string bzHlt  = KbFilterZiweiContent(KbExtractPalaceSection(bzFullContent, "疾厄宮"), KbGetPalaceStarsSet(bzPalaces, "疾厄"), bzChartStars);

                    // 4.2/4.3 雙星或單星入宮（6/7/8 docx）
                    string bzDescMing = await KbQueryStarInPalace(bzPalaces, "命宮");
                    string bzDescOff  = await KbQueryStarInPalace(bzPalaces, "官祿宮");
                    string bzDescWlt  = await KbQueryStarInPalace(bzPalaces, "財帛宮");
                    string bzDescSps  = await KbQueryStarInPalace(bzPalaces, "夫妻宮");
                    string bzDescHlt  = await KbQueryStarInPalace(bzPalaces, "疾厄宮");
                    string bzDoubleDescMing = await KbQueryDoubleStarInPalace(bzPalaces, "命宮");
                    string bzDoubleDescOff  = await KbQueryDoubleStarInPalace(bzPalaces, "官祿宮");
                    string bzDoubleDescWlt  = await KbQueryDoubleStarInPalace(bzPalaces, "財帛宮");
                    string bzDoubleDescSps  = await KbQueryDoubleStarInPalace(bzPalaces, "夫妻宮");
                    string bzDoubleDescHlt  = await KbQueryDoubleStarInPalace(bzPalaces, "疾厄宮");

                    // 4.4 副星/小星入宮（吉星 9.docx + 煞星 4.docx）
                    string bzMinorDescMing = await KbQueryMinorStarsInPalace(bzPalaces, "命宮");
                    string bzMinorDescOff  = await KbQueryMinorStarsInPalace(bzPalaces, "官祿宮");
                    string bzMinorDescWlt  = await KbQueryMinorStarsInPalace(bzPalaces, "財帛宮");
                    string bzMinorDescSps  = await KbQueryMinorStarsInPalace(bzPalaces, "夫妻宮");
                    string bzMinorDescHlt  = await KbQueryMinorStarsInPalace(bzPalaces, "疾厄宮");

                    string bzOffStars = KbGetPalaceStars(bzPalaces, "官祿");
                    string bzWltStars = KbGetPalaceStars(bzPalaces, "財帛");
                    string bzSpsStars = KbGetPalaceStars(bzPalaces, "夫妻");
                    string bzHltStars = KbGetPalaceStars(bzPalaces, "疾厄");

                    // 其餘七宮（兄弟、子女、遷移、交友、田宅、福德、父母）
                    string bzBrtStars = KbGetPalaceStars(bzPalaces, "兄弟");
                    string bzChdStars = KbGetPalaceStars(bzPalaces, "子女");
                    string bzTrvStars = KbGetPalaceStars(bzPalaces, "遷移");
                    string bzFrdStars = KbGetPalaceStars(bzPalaces, "奴僕"); // JSON 存為奴僕宮
                    string bzHseStars = KbGetPalaceStars(bzPalaces, "田宅");
                    string bzFrtStars = KbGetPalaceStars(bzPalaces, "福德");
                    string bzPrtStars = KbGetPalaceStars(bzPalaces, "父母");
                    string bzDescBrt = await KbQueryStarInPalace(bzPalaces, "兄弟宮");
                    string bzDescChd = await KbQueryStarInPalace(bzPalaces, "子女宮");
                    string bzDescTrv = await KbQueryStarInPalace(bzPalaces, "遷移宮");
                    string bzDescFrd = await KbQueryStarInPalace(bzPalaces, "奴僕宮"); // JSON 存為奴僕宮
                    string bzDescHse = await KbQueryStarInPalace(bzPalaces, "田宅宮");
                    string bzDescFrt = await KbQueryStarInPalace(bzPalaces, "福德宮");
                    string bzDescPrt = await KbQueryStarInPalace(bzPalaces, "父母宮");

                    // 宮位顯示名稱（奴僕宮→交友宮，官祿宮保持）
                    string BzDisplayPalace(string p) => p switch
                    {
                        "奴僕宮" => "交友宮", "官祿宮" => "事業宮",  _ => p
                    };

                    // Helper to append 四化 for any palace (使用 KbPalaceSame 做模糊比對)
                    void BzAppendSiHua(System.Text.StringBuilder s, string palaceName)
                    {
                        if (!string.IsNullOrEmpty(bzSiHuaLu) && KbPalaceSame(bzSiHuaLuPalace, palaceName))
                            s.AppendLine($"【化祿加持】{bzSiHuaLu}");
                        if (!string.IsNullOrEmpty(bzSiHuaQuan) && KbPalaceSame(bzSiHuaQuanPalace, palaceName))
                            s.AppendLine($"【化權加持】{bzSiHuaQuan}");
                        if (!string.IsNullOrEmpty(bzSiHuaKe) && KbPalaceSame(bzSiHuaKePalace, palaceName))
                            s.AppendLine($"【化科加持】{bzSiHuaKe}");
                        if (!string.IsNullOrEmpty(bzSiHuaJi) && KbPalaceSame(bzSiHuaJiPalace, palaceName))
                            s.AppendLine($"【化忌警示】{bzSiHuaJi}");
                    }

                    var bzSb = new System.Text.StringBuilder();
                    bzSb.AppendLine("=================================================================");
                    bzSb.AppendLine("【第十一章：紫微星盤鑑定】");
                    bzSb.AppendLine("=================================================================");
                    if (!string.IsNullOrEmpty(bzMingGongStars)) bzSb.AppendLine($"命宮主星：{bzMingGongStars}");
                    if (!string.IsNullOrEmpty(bzMingZhu))       bzSb.AppendLine($"命主：{bzMingZhu}");
                    if (!string.IsNullOrEmpty(bzShenZhu))       bzSb.AppendLine($"身主：{bzShenZhu}");
                    if (!string.IsNullOrEmpty(bzWuXingJu))      bzSb.AppendLine($"五行局：{bzWuXingJu}");
                    bzSb.AppendLine();

                    // 命格格局（4.1）
                    if (!string.IsNullOrEmpty(bzGeJu))
                    {
                        bzSb.AppendLine("【命格格局】");
                        bzSb.AppendLine(bzGeJu.TrimEnd());
                        bzSb.AppendLine();
                    }

                    // 終身影響提醒（先天四化白話說明）
                    {
                        var siHuaSb = new System.Text.StringBuilder();
                        if (!string.IsNullOrEmpty(bzSiHuaLu))   siHuaSb.AppendLine($"一生財祿之源在{BzDisplayPalace(bzSiHuaLuPalace)}：{bzSiHuaLu}");
                        if (!string.IsNullOrEmpty(bzSiHuaQuan)) siHuaSb.AppendLine($"一生主導力量在{BzDisplayPalace(bzSiHuaQuanPalace)}：{bzSiHuaQuan}");
                        if (!string.IsNullOrEmpty(bzSiHuaKe))   siHuaSb.AppendLine($"一生名聲科甲在{BzDisplayPalace(bzSiHuaKePalace)}：{bzSiHuaKe}");
                        if (!string.IsNullOrEmpty(bzSiHuaJi))   siHuaSb.AppendLine($"一生需謹慎之處在{BzDisplayPalace(bzSiHuaJiPalace)}（化忌）：{bzSiHuaJi}");
                        if (siHuaSb.Length > 0)
                        {
                            bzSb.AppendLine("【終身影響提醒】");
                            bzSb.Append(siHuaSb);
                            bzSb.AppendLine();
                        }
                    }

                    // 命宮
                    bzSb.AppendLine($"--- 命宮（{bzMingGongStars}）---");
                    if (!string.IsNullOrEmpty(bzDoubleDescMing))
                        bzSb.AppendLine($"【雙星論斷】{bzDoubleDescMing}");
                    else if (!string.IsNullOrEmpty(bzDescMing))
                        bzSb.AppendLine($"【主星星情】{bzDescMing}");
                    if (!string.IsNullOrEmpty(bzMing)) bzSb.AppendLine(bzMing);
                    if (!string.IsNullOrEmpty(bzMinorDescMing)) bzSb.AppendLine($"【輔星加臨】{bzMinorDescMing}");
                    BzAppendSiHua(bzSb, "命宮");
                    bzSb.AppendLine();

                    // 兄弟宮
                    bzSb.AppendLine($"--- 兄弟宮（{bzBrtStars}）---");
                    if (!string.IsNullOrEmpty(bzDescBrt)) bzSb.AppendLine($"【主星星情】{bzDescBrt}");
                    BzAppendSiHua(bzSb, "兄弟宮");
                    bzSb.AppendLine();

                    // 夫妻宮
                    bzSb.AppendLine($"--- 夫妻宮（{bzSpsStars}）---");
                    if (!string.IsNullOrEmpty(bzDoubleDescSps))
                        bzSb.AppendLine($"【雙星論斷】{bzDoubleDescSps}");
                    else if (!string.IsNullOrEmpty(bzDescSps))
                        bzSb.AppendLine($"【主星星情】{bzDescSps}");
                    if (!string.IsNullOrEmpty(bzSps)) bzSb.AppendLine(bzSps);
                    if (!string.IsNullOrEmpty(bzMinorDescSps)) bzSb.AppendLine($"【輔星加臨】{bzMinorDescSps}");
                    BzAppendSiHua(bzSb, "夫妻宮");
                    bzSb.AppendLine();

                    // 子女宮
                    bzSb.AppendLine($"--- 子女宮（{bzChdStars}）---");
                    if (!string.IsNullOrEmpty(bzDescChd)) bzSb.AppendLine($"【主星星情】{bzDescChd}");
                    BzAppendSiHua(bzSb, "子女宮");
                    bzSb.AppendLine();

                    // 財帛宮
                    bzSb.AppendLine($"--- 財帛宮（{bzWltStars}）---");
                    if (!string.IsNullOrEmpty(bzDoubleDescWlt))
                        bzSb.AppendLine($"【雙星論斷】{bzDoubleDescWlt}");
                    else if (!string.IsNullOrEmpty(bzDescWlt))
                        bzSb.AppendLine($"【主星星情】{bzDescWlt}");
                    if (!string.IsNullOrEmpty(bzWlt)) bzSb.AppendLine(bzWlt);
                    if (!string.IsNullOrEmpty(bzMinorDescWlt)) bzSb.AppendLine($"【輔星加臨】{bzMinorDescWlt}");
                    BzAppendSiHua(bzSb, "財帛宮");
                    bzSb.AppendLine();

                    // 疾厄宮
                    bzSb.AppendLine($"--- 疾厄宮（{bzHltStars}）---");
                    if (!string.IsNullOrEmpty(bzDoubleDescHlt))
                        bzSb.AppendLine($"【雙星論斷】{bzDoubleDescHlt}");
                    else if (!string.IsNullOrEmpty(bzDescHlt))
                        bzSb.AppendLine($"【主星星情】{bzDescHlt}");
                    if (!string.IsNullOrEmpty(bzHlt)) bzSb.AppendLine(bzHlt);
                    if (!string.IsNullOrEmpty(bzMinorDescHlt)) bzSb.AppendLine($"【輔星加臨】{bzMinorDescHlt}");
                    BzAppendSiHua(bzSb, "疾厄宮");
                    bzSb.AppendLine();

                    // 遷移宮
                    bzSb.AppendLine($"--- 遷移宮（{bzTrvStars}）---");
                    if (!string.IsNullOrEmpty(bzDescTrv)) bzSb.AppendLine($"【主星星情】{bzDescTrv}");
                    BzAppendSiHua(bzSb, "遷移宮");
                    bzSb.AppendLine();

                    // 交友宮（僕役宮）
                    bzSb.AppendLine($"--- 交友宮（{bzFrdStars}）---");
                    if (!string.IsNullOrEmpty(bzDescFrd)) bzSb.AppendLine($"【主星星情】{bzDescFrd}");
                    BzAppendSiHua(bzSb, "交友宮");
                    bzSb.AppendLine();

                    // 官祿宮
                    bzSb.AppendLine($"--- 官祿宮（{bzOffStars}）---");
                    if (!string.IsNullOrEmpty(bzDoubleDescOff))
                        bzSb.AppendLine($"【雙星論斷】{bzDoubleDescOff}");
                    else if (!string.IsNullOrEmpty(bzDescOff))
                        bzSb.AppendLine($"【主星星情】{bzDescOff}");
                    if (!string.IsNullOrEmpty(bzOff)) bzSb.AppendLine(bzOff);
                    if (!string.IsNullOrEmpty(bzMinorDescOff)) bzSb.AppendLine($"【輔星加臨】{bzMinorDescOff}");
                    BzAppendSiHua(bzSb, "官祿宮");
                    bzSb.AppendLine();

                    // 田宅宮
                    bzSb.AppendLine($"--- 田宅宮（{bzHseStars}）---");
                    if (!string.IsNullOrEmpty(bzDescHse)) bzSb.AppendLine($"【主星星情】{bzDescHse}");
                    BzAppendSiHua(bzSb, "田宅宮");
                    bzSb.AppendLine();

                    // 福德宮
                    bzSb.AppendLine($"--- 福德宮（{bzFrtStars}）---");
                    if (!string.IsNullOrEmpty(bzDescFrt)) bzSb.AppendLine($"【主星星情】{bzDescFrt}");
                    BzAppendSiHua(bzSb, "福德宮");
                    bzSb.AppendLine();

                    // 父母宮
                    bzSb.AppendLine($"--- 父母宮（{bzPrtStars}）---");
                    if (!string.IsNullOrEmpty(bzDescPrt)) bzSb.AppendLine($"【主星星情】{bzDescPrt}");
                    BzAppendSiHua(bzSb, "父母宮");
                    bzSb.AppendLine();

                    report = report.TrimEnd() + Environment.NewLine + bzSb.ToString();
                }

                // === 九星氣學加成 ===
                string bzNsSection = await NsBuildBirthSection(
                    user.BirthYear ?? DateTime.Today.Year - 30,
                    user.BirthMonth ?? 1,
                    user.BirthDay ?? 1,
                    user.BirthHour ?? 0,
                    user.BirthGender ?? 1);
                if (!string.IsNullOrEmpty(bzNsSection)) report += bzNsSection;

                var cycleData = scored.Select(c => new {
                    stem = c.stem, branch = c.branch, liuShen = c.liuShen,
                    startAge = c.startAge, endAge = c.endAge,
                    score = c.score, level = c.level
                }).ToList();

                var baziTable = new {
                    pillars = new[] {
                        new { label = "年", stem = yStem, branch = yBranch, stemSS = yStemSS,
                              naYin = LfPillarNaYin(yearP), hiddenPairs = LfPillarHiddenPairs(yearP) },
                        new { label = "月", stem = mStem, branch = mBranch, stemSS = mStemSS,
                              naYin = LfPillarNaYin(monthP), hiddenPairs = LfPillarHiddenPairs(monthP) },
                        new { label = "日", stem = dStem, branch = dBranch, stemSS = "元神",
                              naYin = LfPillarNaYin(dayP), hiddenPairs = LfPillarHiddenPairs(dayP) },
                        new { label = "時", stem = hStem, branch = hBranch, stemSS = hStemSS,
                              naYin = LfPillarNaYin(timeP), hiddenPairs = LfPillarHiddenPairs(timeP) },
                    }
                };

                string bzTuneElem = season == "冬" ? "火" : season == "夏" ? "水" : "";
                string bzJiYongElem = LfElemOvercomeBy.GetValueOrDefault(yongShenElem, "");
                string BzCls(string elem) {
                    if (elem == jiShenElem) return "X";
                    if (elem == yongShenElem || elem == fuYiElem || (!string.IsNullOrEmpty(bzTuneElem) && elem == bzTuneElem)) return "○";
                    if (elem == bzJiYongElem && bzJiYongElem != jiShenElem) return "△忌";
                    return "△";
                }
                string[] bzAllStems = { "甲","乙","丙","丁","戊","己","庚","辛","壬","癸" };
                string[] bzAllBrs   = { "子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥" };
                var yongJiTable = new {
                    stems = bzAllStems.Select(s => new {
                        stem = s, elem = KbStemToElement(s),
                        shiShen = LfStemShiShen(s, dStem), cls = BzCls(KbStemToElement(s))
                    }).ToArray(),
                    branches = bzAllBrs.Select(br => {
                        string brElem = LfBranchHiddenRatio.TryGetValue(br, out var bh) && bh.Count > 0 ? KbStemToElement(bh[0].stem) : "-";
                        string brMs = LfBranchHiddenRatio.TryGetValue(br, out var bh2) && bh2.Count > 0 ? bh2[0].stem : "";
                        string brSS = !string.IsNullOrEmpty(brMs) ? LfStemShiShen(brMs, dStem) : "-";
                        return new { branch = br, elem = brElem, shiShen = brSS, cls = brElem != "-" ? BzCls(brElem) : "-", inChart = branches.Contains(br) };
                    }).ToArray()
                };

                if (!bzIsAdmin) await RecordSubscriptionClaim(user.Id, bzSubId, "BOOK_BAZI");
                await SaveUserReportAsync(user.Id, "bazi-ziwei", "玉洞子八字紫微命書", report,
                    new { birthYear = user.BirthYear, birthMonth = user.BirthMonth, birthDay = user.BirthDay, gender = user.BirthGender });
                return Ok(new { result = report, luckCycles = cycleData, baziTable, yongJiTable, remainingPoints = user.Points });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "八字紫微命書失敗 User={User}", identity);
                return StatusCode(500, new { error = "八字紫微命書生成失敗，請稍後再試", details = ex.Message });
            }
        }

        [HttpGet("analyze-yudongzi")]
        [Authorize]
        public async Task<IActionResult> GetYudongziAnalysis([FromQuery] string? personName = null)
        {
            var identity = User.FindFirstValue(ClaimTypes.Email)
                         ?? User.FindFirstValue(ClaimTypes.Name)
                         ?? User.FindFirst("unique_name")?.Value;
            if (string.IsNullOrEmpty(identity))
                return Unauthorized(new { error = "請重新登入" });

            var user = await _context.Users.FirstOrDefaultAsync(u => u.UserName == identity || u.Email == identity);
            if (user == null) return BadRequest(new { error = "找不到用戶" });

            bool ydIsAdmin = string.Equals(user.Email, _config["Admin:Email"], StringComparison.OrdinalIgnoreCase);
            int ydSubId = -1;
            if (!ydIsAdmin)
            {
                var (ydOk, ydErr, ydSubIdVal) = await CheckSubscriptionQuota(user.Id, "BOOK_VIP");
                if (!ydOk) return BadRequest(new { error = ydErr });
                ydSubId = ydSubIdVal;
            }

            var userChart = await _context.UserCharts.FirstOrDefaultAsync(c => c.UserId == user.Id);
            if (userChart == null || string.IsNullOrEmpty(userChart.ChartJson))
                return BadRequest(new { error = "no_chart" });

            try
            {
                var root = JsonDocument.Parse(userChart.ChartJson).RootElement;
                if (!root.TryGetProperty("bazi", out var bazi) && !root.TryGetProperty("baziInfo", out bazi))
                    return BadRequest(new { error = "命盤資料格式錯誤" });

                string lunarRawYdz = root.TryGetProperty("lunarBirthDate", out var lbYdzEl) ? lbYdzEl.GetString() ?? "" : "";
                int lunarMonthYdz = LfParseLunarMonth(lunarRawYdz);

                var yearP  = LfGetPillar(bazi, "yearPillar");
                var monthP = LfGetPillar(bazi, "monthPillar");
                var dayP   = LfGetPillar(bazi, "dayPillar");
                var timeP  = LfGetPillar(bazi, "timePillar");

                string yStem = LfPillarStem(yearP);   string yBranch = LfPillarBranch(yearP);
                string mStem = LfPillarStem(monthP);  string mBranch = LfPillarBranch(monthP);
                string dStem = LfPillarStem(dayP);    string dBranch = LfPillarBranch(dayP);
                string hStem = LfPillarStem(timeP);   string hBranch = LfPillarBranch(timeP);

                string yStemSS   = LfPillarStemSS(yearP);
                string mStemSS   = LfPillarStemSS(monthP);
                string hStemSS   = LfPillarStemSS(timeP);
                string yBranchSS = LfPillarBranchMainSS(yearP);
                string mBranchSS = LfPillarBranchMainSS(monthP);
                string dBranchSS = LfPillarBranchMainSS(dayP);
                string hBranchSS = LfPillarBranchMainSS(timeP);

                int birthYear = user.BirthYear ?? (DateTime.Today.Year - 30);
                int gender    = user.BirthGender ?? 1;
                string ydzUserName = !string.IsNullOrEmpty(personName) ? personName : (user.Name ?? "");
                var luckCycles = LfExtractLuckCycles(root);

                string dmElem  = KbStemToElement(dStem);
                var branches   = new[] { yBranch, mBranch, dBranch, hBranch };
                var wuXing     = LfCalcWuXingMatrix(yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch);
                double bodyPct = LfGetBodyStrengthPct(dmElem, wuXing);
                string bodyLabel = LfGetBodyStrengthLabel(bodyPct);
                string season  = LfGetSeason(mBranch);
                string seaLabel = LfGetSeasonLabel(mBranch);

                var (pattern, yongShenElem, fuYiElem, yongReason, tiaoHouElem) = LfDetectGeJuAndYongShen(
                    yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch,
                    dmElem, wuXing, bodyPct, season);
                string jiShenElem = LfGetJiShenElem(yongShenElem, dmElem, bodyPct, pattern);

                var chartStems2 = new[] { yStem, mStem, dStem, hStem };
                var scored = luckCycles.Select(lc =>
                {
                    int sc = LfCalcLuckScore(lc.stem, lc.branch, pattern, yongShenElem, fuYiElem, jiShenElem,
                        dmElem, bodyPct > 50, tiaoHouElem, season, branches, chartStems2, dStem);
                    return (lc.stem, lc.branch, lc.liuShen, lc.startAge, lc.endAge, score: sc, level: LfLuckLevel(sc));
                }).ToList();

                // Query day pillar KB
                var dayPillarKb = await _context.BaziDayPillarReadings
                    .FirstOrDefaultAsync(r => r.DayPillar == dStem + dBranch);

                // 中原盲派直斷規則
                var zhongyuanRules = await _context.BaziDirectRules
                    .Where(r => r.Chapter >= 4)
                    .OrderBy(r => r.Chapter).ThenBy(r => r.Section).ThenBy(r => r.SortOrder)
                    .ToListAsync();

                // NaYin strings
                string yNaYin = LfPillarNaYin(yearP);
                string mNaYin = LfPillarNaYin(monthP);
                string dNaYin = LfPillarNaYin(dayP);
                string hNaYin = LfPillarNaYin(timeP);

                // Ziwei data
                bool hasZiwei = root.TryGetProperty("palaces", out var palacesYdz)
                    && palacesYdz.ValueKind == JsonValueKind.Array;
                string mingGongStarsYdz = userChart.MingGongMainStars ?? "";
                string mingZhuYdz = root.TryGetProperty("mingZhu", out var mzYdz) ? mzYdz.GetString() ?? "" : "";
                string shenZhuYdz = root.TryGetProperty("shenZhu", out var szYdz) ? szYdz.GetString() ?? "" : "";
                string wuXingJuTextYdz = root.TryGetProperty("wuXingJuText", out var wxYdz) ? wxYdz.GetString() ?? "" : "";

                string ziweiFullContentYdz = "";
                var chartStarsYdz = new HashSet<string>();
                var siHuaYdz = new Dictionary<string, (string pal, string txt)>();
                string starDescMingYdz = "";
                string ziweiMingYdz = "", ziweiOffYdz = "", offStarsYdz = "";
                string ziweiWltYdz = "", wltStarsYdz = "";
                string ziweiSpsYdz = "", spsStarsYdz = "";
                string ziweiHltYdz = "", hltStarsYdz = "";
                string starDescOffYdz = "", starDescWltYdz = "", starDescSpsYdz = "", starDescHltYdz = "";
                string ziweiParStarYdz = "", ziweiParYdz = "", ziweiCldStarYdz = "", ziweiCldYdz = "";
                string nianSiHuaXingYdz = "";
                string siHuaLuPalaceYdz = "", siHuaLuYdz = "";
                string siHuaQuanPalaceYdz = "", siHuaQuanYdz = "";
                string siHuaKePalaceYdz = "", siHuaKeYdz = "";
                string siHuaJiPalaceYdz = "", siHuaJiYdz = "";

                // Fallback: if saved chart has no palaces, recalculate Ziwei from user birth data
                if (!hasZiwei && user.BirthYear.HasValue && user.BirthMonth.HasValue && user.BirthDay.HasValue)
                {
                    try
                    {
                        var ziweiReq = new AstrologyRequest(
                            user.BirthYear.Value, user.BirthMonth.Value, user.BirthDay.Value,
                            user.BirthHour ?? 0, user.BirthMinute ?? 0,
                            user.BirthGender ?? 1, user.ChartName ?? user.UserName ?? "",
                            user.DateType ?? "solar");
                        var freshChart = await _astrologyService.CalculateChartAsync(ziweiReq);
                        var ccOpts = new System.Text.Json.JsonSerializerOptions { PropertyNamingPolicy = System.Text.Json.JsonNamingPolicy.CamelCase };
                        string freshJson = System.Text.Json.JsonSerializer.Serialize(freshChart, ccOpts);
                        var freshRoot = System.Text.Json.JsonDocument.Parse(freshJson).RootElement;
                        hasZiwei = freshRoot.TryGetProperty("palaces", out palacesYdz) && palacesYdz.ValueKind == System.Text.Json.JsonValueKind.Array;
                        if (hasZiwei)
                        {
                            if (freshRoot.TryGetProperty("mingZhu", out var mzFresh)) mingZhuYdz = mzFresh.GetString() ?? "";
                            if (freshRoot.TryGetProperty("shenZhu", out var szFresh)) shenZhuYdz = szFresh.GetString() ?? "";
                            if (freshRoot.TryGetProperty("wuXingJuText", out var wxFresh)) wuXingJuTextYdz = wxFresh.GetString() ?? "";
                            userChart.ChartJson = freshJson;
                            userChart.UpdatedAt = DateTime.UtcNow;
                            await _context.SaveChangesAsync();
                        }
                    }
                    catch (Exception exFresh) { _logger.LogWarning(exFresh, "Yudongzi Ziwei recalc fallback failed"); }
                }

                if (hasZiwei)
                {
                    string zPosYdz = KbGetZiweiPosition(palacesYdz);
                    chartStarsYdz = KbGetAllChartStars(palacesYdz);
                    ziweiFullContentYdz = await KbZiweiFullQuery(palacesYdz, zPosYdz);

                    ziweiMingYdz = KbFilterZiweiContent(KbExtractPalaceSection(ziweiFullContentYdz, "命宮"),   KbGetPalaceStarsSet(palacesYdz, "命宮"),  chartStarsYdz);
                    ziweiOffYdz  = KbFilterZiweiContent(KbExtractPalaceSection(ziweiFullContentYdz, "事業宮"), KbGetPalaceStarsSet(palacesYdz, "官祿"),  chartStarsYdz);
                    ziweiWltYdz  = KbFilterZiweiContent(KbExtractPalaceSection(ziweiFullContentYdz, "財帛宮"), KbGetPalaceStarsSet(palacesYdz, "財帛"),  chartStarsYdz);
                    ziweiSpsYdz  = KbFilterZiweiContent(KbExtractPalaceSection(ziweiFullContentYdz, "夫妻宮"), KbGetPalaceStarsSet(palacesYdz, "夫妻"),  chartStarsYdz);
                    ziweiHltYdz  = KbFilterZiweiContent(KbExtractPalaceSection(ziweiFullContentYdz, "疾厄宮"), KbGetPalaceStarsSet(palacesYdz, "疾厄"),  chartStarsYdz);
                    starDescMingYdz = await KbQueryStarInPalace(palacesYdz, "命宮");
                    offStarsYdz  = KbGetPalaceStars(palacesYdz, "官祿");
                    wltStarsYdz  = KbGetPalaceStars(palacesYdz, "財帛");
                    spsStarsYdz  = KbGetPalaceStars(palacesYdz, "夫妻");
                    hltStarsYdz  = KbGetPalaceStars(palacesYdz, "疾厄");
                    starDescOffYdz = await KbQueryStarInPalace(palacesYdz, "官祿宮");
                    starDescWltYdz = await KbQueryStarInPalace(palacesYdz, "財帛宮");
                    starDescSpsYdz = await KbQueryStarInPalace(palacesYdz, "夫妻宮");
                    starDescHltYdz = await KbQueryStarInPalace(palacesYdz, "疾厄宮");
                    ziweiParStarYdz = KbGetPalaceStars(palacesYdz, "父母");
                    ziweiParYdz = KbFilterZiweiContent(KbExtractPalaceSection(ziweiFullContentYdz, "父母宮"), KbGetPalaceStarsSet(palacesYdz, "父母"), chartStarsYdz);
                    ziweiCldStarYdz = KbGetPalaceStars(palacesYdz, "子女");
                    ziweiCldYdz = KbFilterZiweiContent(KbExtractPalaceSection(ziweiFullContentYdz, "子女宮"), KbGetPalaceStarsSet(palacesYdz, "子女"), chartStarsYdz);

                    // 宮位四化：6宮 × 4化（命宮/官祿宮/財帛宮/夫妻宮/疾厄宮/遷移宮）
                    var (mLuP, mLuC)   = await KbGongWeiSiHuaQuery(palacesYdz, "命宮",   "化祿");
                    var (mQuanP,mQuanC)= await KbGongWeiSiHuaQuery(palacesYdz, "命宮",   "化權");
                    var (mKeP, mKeC)   = await KbGongWeiSiHuaQuery(palacesYdz, "命宮",   "化科");
                    var (mJiP, mJiC)   = await KbGongWeiSiHuaQuery(palacesYdz, "命宮",   "化忌");
                    var (oLuP, oLuC)   = await KbGongWeiSiHuaQuery(palacesYdz, "官祿宮", "化祿");
                    var (oQuanP,oQuanC)= await KbGongWeiSiHuaQuery(palacesYdz, "官祿宮", "化權");
                    var (oKeP, oKeC)   = await KbGongWeiSiHuaQuery(palacesYdz, "官祿宮", "化科");
                    var (oJiP, oJiC)   = await KbGongWeiSiHuaQuery(palacesYdz, "官祿宮", "化忌");
                    var (wLuP, wLuC)   = await KbGongWeiSiHuaQuery(palacesYdz, "財帛宮", "化祿");
                    var (wQuanP,wQuanC)= await KbGongWeiSiHuaQuery(palacesYdz, "財帛宮", "化權");
                    var (wKeP, wKeC)   = await KbGongWeiSiHuaQuery(palacesYdz, "財帛宮", "化科");
                    var (wJiP, wJiC)   = await KbGongWeiSiHuaQuery(palacesYdz, "財帛宮", "化忌");
                    var (sLuP, sLuC)   = await KbGongWeiSiHuaQuery(palacesYdz, "夫妻宮", "化祿");
                    var (sQuanP,sQuanC)= await KbGongWeiSiHuaQuery(palacesYdz, "夫妻宮", "化權");
                    var (sKeP, sKeC)   = await KbGongWeiSiHuaQuery(palacesYdz, "夫妻宮", "化科");
                    var (sJiP, sJiC)   = await KbGongWeiSiHuaQuery(palacesYdz, "夫妻宮", "化忌");
                    var (hLuP, hLuC)   = await KbGongWeiSiHuaQuery(palacesYdz, "疾厄宮", "化祿");
                    var (hQuanP,hQuanC)= await KbGongWeiSiHuaQuery(palacesYdz, "疾厄宮", "化權");
                    var (hKeP, hKeC)   = await KbGongWeiSiHuaQuery(palacesYdz, "疾厄宮", "化科");
                    var (hJiP, hJiC)   = await KbGongWeiSiHuaQuery(palacesYdz, "疾厄宮", "化忌");
                    var (qQuanP,qQuanC)= await KbGongWeiSiHuaQuery(palacesYdz, "遷移宮", "化權");
                    var (qKeP, qKeC)   = await KbGongWeiSiHuaQuery(palacesYdz, "遷移宮", "化科");
                    var (qJiP, qJiC)   = await KbGongWeiSiHuaQuery(palacesYdz, "遷移宮", "化忌");
                    siHuaYdz["命宮化祿"] = (mLuP,   mLuC);   siHuaYdz["命宮化權"] = (mQuanP, mQuanC);
                    siHuaYdz["命宮化科"] = (mKeP,   mKeC);   siHuaYdz["命宮化忌"] = (mJiP,   mJiC);
                    siHuaYdz["官祿化祿"] = (oLuP,   oLuC);   siHuaYdz["官祿化權"] = (oQuanP, oQuanC);
                    siHuaYdz["官祿化科"] = (oKeP,   oKeC);   siHuaYdz["官祿化忌"] = (oJiP,   oJiC);
                    siHuaYdz["財帛化祿"] = (wLuP,   wLuC);   siHuaYdz["財帛化權"] = (wQuanP, wQuanC);
                    siHuaYdz["財帛化科"] = (wKeP,   wKeC);   siHuaYdz["財帛化忌"] = (wJiP,   wJiC);
                    siHuaYdz["夫妻化祿"] = (sLuP,   sLuC);   siHuaYdz["夫妻化權"] = (sQuanP, sQuanC);
                    siHuaYdz["夫妻化科"] = (sKeP,   sKeC);   siHuaYdz["夫妻化忌"] = (sJiP,   sJiC);
                    siHuaYdz["疾厄化祿"] = (hLuP,   hLuC);   siHuaYdz["疾厄化權"] = (hQuanP, hQuanC);
                    siHuaYdz["疾厄化科"] = (hKeP,   hKeC);   siHuaYdz["疾厄化忌"] = (hJiP,   hJiC);
                    siHuaYdz["遷移化權"] = (qQuanP, qQuanC); siHuaYdz["遷移化科"] = (qKeP,   qKeC);
                    siHuaYdz["遷移化忌"] = (qJiP,   qJiC);
                    // 先天四化（年干）
                    nianSiHuaXingYdz = await KbQuery($"SELECT COALESCE(\"ResultText\",'') AS \"Value\" FROM \"FortuneRules\" WHERE \"SourceFile\"='四化干性.docx' AND \"Title\" LIKE '{yStem}年干%' LIMIT 1");
                    siHuaLuPalaceYdz   = KbGetSiHuaPalace(yStem, "化祿", palacesYdz);
                    siHuaQuanPalaceYdz = KbGetSiHuaPalace(yStem, "化權", palacesYdz);
                    siHuaKePalaceYdz   = KbGetSiHuaPalace(yStem, "化科", palacesYdz);
                    siHuaJiPalaceYdz   = KbGetSiHuaPalace(yStem, "化忌", palacesYdz);
                    siHuaLuYdz   = await KbSiHuaQuery(yStem, "化祿", palacesYdz);
                    siHuaQuanYdz = await KbSiHuaQuery(yStem, "化權", palacesYdz);
                    siHuaKeYdz   = await KbSiHuaQuery(yStem, "化科", palacesYdz);
                    siHuaJiYdz   = await KbSiHuaQuery(yStem, "化忌", palacesYdz);
                }

                // 紫微格局偵測 + 描述查詢
                string ziweiGeJuYdz = "";
                if (hasZiwei)
                {
                    string mingBrYdz = KbGetPalaceBranch(palacesYdz, "命宮");
                    var gjList = LfDetectZiweiGeJu(mingGongStarsYdz, mingBrYdz, chartStarsYdz,
                        siHuaLuPalaceYdz, siHuaQuanPalaceYdz, siHuaKePalaceYdz, palacesYdz);
                    var gjSb = new StringBuilder();
                    foreach (var gj in gjList)
                    {
                        string desc = await KbQuery($"SELECT COALESCE(\"ResultText\",'') AS \"Value\" FROM \"FortuneRules\" WHERE \"SourceFile\"='紫微格局說明.docx' AND \"Title\"='{gj}' LIMIT 1");
                        if (!string.IsNullOrEmpty(desc)) { gjSb.AppendLine($"【{gj}】"); gjSb.AppendLine(desc); gjSb.AppendLine(); }
                    }
                    ziweiGeJuYdz = gjSb.ToString();
                }

                // 雙星組合 + 輔星入宮（5宮）
                var doubleDescsYdz = new Dictionary<string, string>();
                var minorDescsYdz  = new Dictionary<string, string>();
                if (hasZiwei)
                {
                    foreach (var kbPal in new[] { "命宮", "官祿宮", "財帛宮", "夫妻宮", "疾厄宮" })
                    {
                        doubleDescsYdz[kbPal] = await KbQueryDoubleStarInPalace(palacesYdz, kbPal);
                        minorDescsYdz[kbPal]  = await KbQueryMinorStarsInPalace(palacesYdz, kbPal);
                    }
                }

                // 十二宮星情特質（Ch.7 用）
                var allPalaceStarDescsYdz = new Dictionary<string, string>();
                if (hasZiwei)
                {
                    allPalaceStarDescsYdz["命宮"]   = starDescMingYdz;
                    allPalaceStarDescsYdz["官祿宮"] = starDescOffYdz;
                    allPalaceStarDescsYdz["財帛宮"] = starDescWltYdz;
                    allPalaceStarDescsYdz["夫妻宮"] = starDescSpsYdz;
                    allPalaceStarDescsYdz["疾厄宮"] = starDescHltYdz;
                    foreach (var pal in new[] { "兄弟宮", "子女宮", "遷移宮", "奴僕宮", "田宅宮", "福德宮", "父母宮" })
                        allPalaceStarDescsYdz[pal] = await KbQueryStarInPalace(palacesYdz, pal);
                }

                string report = LfBuildYudongziReportV2(
                    yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch,
                    yStemSS, mStemSS, hStemSS, yBranchSS, mBranchSS, dBranchSS, hBranchSS,
                    yNaYin, mNaYin, dNaYin, hNaYin,
                    dmElem, wuXing, bodyPct, bodyLabel, season, seaLabel,
                    pattern, yongShenElem, fuYiElem, yongReason, jiShenElem,
                    scored, gender, birthYear, user.BirthMonth, user.BirthDay, user.BirthHour, user.BirthMinute, lunarMonthYdz,
                    hasZiwei, palacesYdz, mingGongStarsYdz, mingZhuYdz, shenZhuYdz, wuXingJuTextYdz,
                    ziweiMingYdz, starDescMingYdz, ziweiFullContentYdz, chartStarsYdz,
                    ziweiOffYdz, offStarsYdz, ziweiWltYdz, wltStarsYdz,
                    ziweiSpsYdz, spsStarsYdz, ziweiHltYdz, hltStarsYdz,
                    siHuaYdz,
                    nianSiHuaXingYdz, yStem,
                    siHuaLuPalaceYdz, siHuaLuYdz,
                    siHuaQuanPalaceYdz, siHuaQuanYdz,
                    siHuaKePalaceYdz, siHuaKeYdz,
                    siHuaJiPalaceYdz, siHuaJiYdz,
                    dayPillarKb, zhongyuanRules, ziweiGeJuYdz,
                    doubleDescsYdz, minorDescsYdz, allPalaceStarDescsYdz,
                    starDescOffYdz, starDescWltYdz, starDescSpsYdz, starDescHltYdz,
                    ziweiParStarYdz, ziweiParYdz, ziweiCldStarYdz, ziweiCldYdz,
                    userName: ydzUserName,
                    calDb: _calendarDb);

                var cycleData = scored.Select(c => new {
                    stem = c.stem, branch = c.branch, liuShen = c.liuShen,
                    startAge = c.startAge, endAge = c.endAge,
                    score = c.score, level = c.level
                }).ToList();

                var baziTable = new {
                    pillars = new[] {
                        new { label = "年", stem = yStem, branch = yBranch, stemSS = yStemSS,
                              naYin = LfPillarNaYin(yearP),
                              hiddenPairs = LfPillarHiddenPairs(yearP) },
                        new { label = "月", stem = mStem, branch = mBranch, stemSS = mStemSS,
                              naYin = LfPillarNaYin(monthP),
                              hiddenPairs = LfPillarHiddenPairs(monthP) },
                        new { label = "日", stem = dStem, branch = dBranch, stemSS = "元神",
                              naYin = LfPillarNaYin(dayP),
                              hiddenPairs = LfPillarHiddenPairs(dayP) },
                        new { label = "時", stem = hStem, branch = hBranch, stemSS = hStemSS,
                              naYin = LfPillarNaYin(timeP),
                              hiddenPairs = LfPillarHiddenPairs(timeP) },
                    }
                };

                string tuneElem = season == "冬" ? "火" : season == "夏" ? "水" : "";
                string jiYongElem = LfElemOvercomeBy.GetValueOrDefault(yongShenElem, "");
                string ClsFor(string elem) {
                    if (elem == jiShenElem) return "X";
                    if (elem == yongShenElem || elem == fuYiElem ||
                        (!string.IsNullOrEmpty(tuneElem) && elem == tuneElem)) return "○";
                    if (elem == jiYongElem && jiYongElem != jiShenElem) return "△忌";
                    return "△";
                }
                string[] allStems = { "甲","乙","丙","丁","戊","己","庚","辛","壬","癸" };
                string[] allBrs   = { "子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥" };
                var yongJiTable = new {
                    stems = allStems.Select(s => new {
                        stem = s,
                        elem = KbStemToElement(s),
                        shiShen = LfStemShiShen(s, dStem),
                        cls = ClsFor(KbStemToElement(s))
                    }).ToArray(),
                    branches = allBrs.Select(br => {
                        string brElem = LfBranchHiddenRatio.TryGetValue(br, out var bh) && bh.Count > 0
                            ? KbStemToElement(bh[0].stem) : "-";
                        string brMs = LfBranchHiddenRatio.TryGetValue(br, out var bh2) && bh2.Count > 0 ? bh2[0].stem : "";
                        string brSS = !string.IsNullOrEmpty(brMs) ? LfStemShiShen(brMs, dStem) : "-";
                        return new {
                            branch = br,
                            elem = brElem,
                            shiShen = brSS,
                            cls = brElem != "-" ? ClsFor(brElem) : "-",
                            inChart = branches.Contains(br)
                        };
                    }).ToArray()
                };

                // 九星氣學加成（純 KB）
                string kbNsSection = await NsBuildBirthSection(
                    user.BirthYear ?? DateTime.Today.Year - 30,
                    user.BirthMonth ?? 1,
                    user.BirthDay ?? 1,
                    user.BirthHour ?? 0,
                    user.BirthGender ?? 1);
                if (!string.IsNullOrEmpty(kbNsSection)) report += kbNsSection;

                if (!ydIsAdmin) await RecordSubscriptionClaim(user.Id, ydSubId, "BOOK_VIP");
                await SaveUserReportAsync(user.Id, "yudongzi", "玉洞子傳家寶典", report,
                    new { birthYear = user.BirthYear, birthMonth = user.BirthMonth, birthDay = user.BirthDay, gender = user.BirthGender });

                return Ok(new { result = report, luckCycles = cycleData, baziTable, yongJiTable, remainingPoints = user.Points,
                    _debug = new { hasZiwei, ziweiWlt = ziweiWltYdz?.Length ?? 0, ziweiOff = ziweiOffYdz?.Length ?? 0, starDescOff = starDescOffYdz?.Length ?? 0, ziweiPar = ziweiParYdz?.Length ?? 0 } });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "玉洞子命書失敗 User={User}", identity);
                return StatusCode(500, new { error = "玉洞子命書生成失敗，請稍後再試", details = ex.Message });
            }
        }

        // === 玉洞子命書 DOCX 匯出 ===

        public class YudongziDocxRequest
        {
            public string? ChartImageBase64 { get; set; }
            public string? ChartJson { get; set; }   // 前端 Astrology/calculate 回傳的完整命盤 JSON
            public string? PersonName { get; set; }  // 客人姓名
        }

        [HttpPost("export-yudongzi-docx")]
        [Authorize]
        public async Task<IActionResult> ExportYudongziDocx([FromBody] YudongziDocxRequest request)
        {
            var identity = User.FindFirstValue(ClaimTypes.Email)
                         ?? User.FindFirstValue(ClaimTypes.Name)
                         ?? User.FindFirst("unique_name")?.Value;
            if (string.IsNullOrEmpty(identity))
                return Unauthorized(new { error = "請重新登入" });

            if (identity != "adam4taiwan@gmail.com")
                return StatusCode(403, new { error = "此功能僅限管理員使用" });

            var user = await _context.Users.FirstOrDefaultAsync(u => u.UserName == identity || u.Email == identity);
            if (user == null) return BadRequest(new { error = "找不到用戶" });

            var userChart = await _context.UserCharts.FirstOrDefaultAsync(c => c.UserId == user.Id);
            if (userChart == null || string.IsNullOrEmpty(userChart.ChartJson))
                return BadRequest(new { error = "no_chart" });

            try
            {
                // === 優先使用前端傳入的命盤 JSON（客人資料），否則 fallback 到 admin 自己的 DB 資料 ===
                string chartJsonToUse = !string.IsNullOrEmpty(request.ChartJson)
                    ? request.ChartJson
                    : userChart.ChartJson;
                var root = JsonDocument.Parse(chartJsonToUse).RootElement;
                if (!root.TryGetProperty("bazi", out var bazi) && !root.TryGetProperty("baziInfo", out bazi))
                    return BadRequest(new { error = "命盤資料格式錯誤" });

                string lunarRawDocx = root.TryGetProperty("lunarBirthDate", out var lbDocxEl) ? lbDocxEl.GetString() ?? "" : "";
                int lunarMonthDocx = LfParseLunarMonth(lunarRawDocx);

                var yearP  = LfGetPillar(bazi, "yearPillar");
                var monthP = LfGetPillar(bazi, "monthPillar");
                var dayP   = LfGetPillar(bazi, "dayPillar");
                var timeP  = LfGetPillar(bazi, "timePillar");

                string yStem = LfPillarStem(yearP);   string yBranch = LfPillarBranch(yearP);
                string mStem = LfPillarStem(monthP);  string mBranch = LfPillarBranch(monthP);
                string dStem = LfPillarStem(dayP);    string dBranch = LfPillarBranch(dayP);
                string hStem = LfPillarStem(timeP);   string hBranch = LfPillarBranch(timeP);

                string yStemSS   = LfPillarStemSS(yearP);
                string mStemSS   = LfPillarStemSS(monthP);
                string hStemSS   = LfPillarStemSS(timeP);
                string yBranchSS = LfPillarBranchMainSS(yearP);
                string mBranchSS = LfPillarBranchMainSS(monthP);
                string dBranchSS = LfPillarBranchMainSS(dayP);
                string hBranchSS = LfPillarBranchMainSS(timeP);

                int birthYear = user.BirthYear ?? (DateTime.Today.Year - 30);
                int gender    = user.BirthGender ?? 1;
                string docxUserName = !string.IsNullOrEmpty(request?.PersonName) ? request.PersonName : (user.Name ?? "");
                var luckCycles = LfExtractLuckCycles(root);

                string dmElem  = KbStemToElement(dStem);
                var branches   = new[] { yBranch, mBranch, dBranch, hBranch };
                var wuXing     = LfCalcWuXingMatrix(yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch);
                double bodyPct = LfGetBodyStrengthPct(dmElem, wuXing);
                string bodyLabel = LfGetBodyStrengthLabel(bodyPct);
                string season  = LfGetSeason(mBranch);
                string seaLabel = LfGetSeasonLabel(mBranch);

                var (pattern, yongShenElem, fuYiElem, yongReason, tiaoHouElem) = LfDetectGeJuAndYongShen(
                    yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch,
                    dmElem, wuXing, bodyPct, season);
                string jiShenElem = LfGetJiShenElem(yongShenElem, dmElem, bodyPct, pattern);

                var chartStems2 = new[] { yStem, mStem, dStem, hStem };
                var scored = luckCycles.Select(lc =>
                {
                    int sc = LfCalcLuckScore(lc.stem, lc.branch, pattern, yongShenElem, fuYiElem, jiShenElem,
                        dmElem, bodyPct > 50, tiaoHouElem, season, branches, chartStems2, dStem);
                    return (lc.stem, lc.branch, lc.liuShen, lc.startAge, lc.endAge, score: sc, level: LfLuckLevel(sc));
                }).ToList();

                var dayPillarKb = await _context.BaziDayPillarReadings
                    .FirstOrDefaultAsync(r => r.DayPillar == dStem + dBranch);
                var zhongyuanRules = await _context.BaziDirectRules
                    .Where(r => r.Chapter >= 4)
                    .OrderBy(r => r.Chapter).ThenBy(r => r.Section).ThenBy(r => r.SortOrder)
                    .ToListAsync();
                string yNaYin = LfPillarNaYin(yearP); string mNaYin = LfPillarNaYin(monthP);
                string dNaYin = LfPillarNaYin(dayP);  string hNaYin = LfPillarNaYin(timeP);

                bool hasZiwei = root.TryGetProperty("palaces", out var palacesYdz) && palacesYdz.ValueKind == JsonValueKind.Array;
                string mingGongStarsYdz = userChart.MingGongMainStars ?? "";
                string mingZhuYdz = ""; string shenZhuYdz = ""; string wuXingJuTextYdz = "";
                if (root.TryGetProperty("mingZhu", out var mzEl)) mingZhuYdz = mzEl.GetString() ?? "";
                if (root.TryGetProperty("shenZhu", out var szEl)) shenZhuYdz = szEl.GetString() ?? "";
                if (root.TryGetProperty("wuXingJuText", out var wxjEl)) wuXingJuTextYdz = wxjEl.GetString() ?? "";

                string zPosYdz = hasZiwei ? KbGetZiweiPosition(palacesYdz) : "";
                var chartStarsYdz = hasZiwei ? KbGetAllChartStars(palacesYdz) : new HashSet<string>();
                string ziweiFullContentYdz = hasZiwei ? await KbZiweiFullQuery(palacesYdz, zPosYdz) : "";
                string ziweiMingYdz = KbFilterZiweiContent(KbExtractPalaceSection(ziweiFullContentYdz, "命宮"), KbGetPalaceStarsSet(palacesYdz, "命宮"), chartStarsYdz);
                string starDescMingYdz = hasZiwei ? await KbQueryStarInPalace(palacesYdz, "命宮") : "";
                string ziweiOffYdz = KbFilterZiweiContent(KbExtractPalaceSection(ziweiFullContentYdz, "事業宮"), KbGetPalaceStarsSet(palacesYdz, "官祿"), chartStarsYdz);
                string offStarsYdz = hasZiwei ? KbGetPalaceStars(palacesYdz, "官祿") : "";
                string ziweiWltYdz = KbFilterZiweiContent(KbExtractPalaceSection(ziweiFullContentYdz, "財帛宮"), KbGetPalaceStarsSet(palacesYdz, "財帛"), chartStarsYdz);
                string wltStarsYdz = hasZiwei ? KbGetPalaceStars(palacesYdz, "財帛") : "";
                string ziweiSpsYdz = KbFilterZiweiContent(KbExtractPalaceSection(ziweiFullContentYdz, "夫妻宮"), KbGetPalaceStarsSet(palacesYdz, "夫妻"), chartStarsYdz);
                string spsStarsYdz = hasZiwei ? KbGetPalaceStars(palacesYdz, "夫妻") : "";
                string ziweiHltYdz = KbFilterZiweiContent(KbExtractPalaceSection(ziweiFullContentYdz, "疾厄宮"), KbGetPalaceStarsSet(palacesYdz, "疾厄"), chartStarsYdz);
                string hltStarsYdz = hasZiwei ? KbGetPalaceStars(palacesYdz, "疾厄") : "";
                string starDescOffYdz = hasZiwei ? await KbQueryStarInPalace(palacesYdz, "官祿宮") : "";
                string starDescWltYdz = hasZiwei ? await KbQueryStarInPalace(palacesYdz, "財帛宮") : "";
                string starDescSpsYdz = hasZiwei ? await KbQueryStarInPalace(palacesYdz, "夫妻宮") : "";
                string starDescHltYdz = hasZiwei ? await KbQueryStarInPalace(palacesYdz, "疾厄宮") : "";
                string ziweiParStarYdz = hasZiwei ? KbGetPalaceStars(palacesYdz, "父母") : "";
                string ziweiParYdz = KbFilterZiweiContent(KbExtractPalaceSection(ziweiFullContentYdz, "父母宮"), KbGetPalaceStarsSet(palacesYdz, "父母"), chartStarsYdz);
                string ziweiCldStarYdz = hasZiwei ? KbGetPalaceStars(palacesYdz, "子女") : "";
                string ziweiCldYdz = KbFilterZiweiContent(KbExtractPalaceSection(ziweiFullContentYdz, "子女宮"), KbGetPalaceStarsSet(palacesYdz, "子女"), chartStarsYdz);
                var siHuaYdz = new Dictionary<string, (string pal, string txt)>();
                string nianSiHuaXingYdz = "";
                string siHuaLuPalaceYdz = "", siHuaLuYdz = "";
                string siHuaQuanPalaceYdz = "", siHuaQuanYdz = "";
                string siHuaKePalaceYdz = "", siHuaKeYdz = "";
                string siHuaJiPalaceYdz = "", siHuaJiYdz = "";
                if (hasZiwei)
                {
                    // 宮位四化：6宮 × 4化（命宮/官祿宮/財帛宮/夫妻宮/疾厄宮/遷移宮）
                    var (mLuP, mLuC)   = await KbGongWeiSiHuaQuery(palacesYdz, "命宮",   "化祿");
                    var (mQuanP,mQuanC)= await KbGongWeiSiHuaQuery(palacesYdz, "命宮",   "化權");
                    var (mKeP, mKeC)   = await KbGongWeiSiHuaQuery(palacesYdz, "命宮",   "化科");
                    var (mJiP, mJiC)   = await KbGongWeiSiHuaQuery(palacesYdz, "命宮",   "化忌");
                    var (oLuP, oLuC)   = await KbGongWeiSiHuaQuery(palacesYdz, "官祿宮", "化祿");
                    var (oQuanP,oQuanC)= await KbGongWeiSiHuaQuery(palacesYdz, "官祿宮", "化權");
                    var (oKeP, oKeC)   = await KbGongWeiSiHuaQuery(palacesYdz, "官祿宮", "化科");
                    var (oJiP, oJiC)   = await KbGongWeiSiHuaQuery(palacesYdz, "官祿宮", "化忌");
                    var (wLuP, wLuC)   = await KbGongWeiSiHuaQuery(palacesYdz, "財帛宮", "化祿");
                    var (wQuanP,wQuanC)= await KbGongWeiSiHuaQuery(palacesYdz, "財帛宮", "化權");
                    var (wKeP, wKeC)   = await KbGongWeiSiHuaQuery(palacesYdz, "財帛宮", "化科");
                    var (wJiP, wJiC)   = await KbGongWeiSiHuaQuery(palacesYdz, "財帛宮", "化忌");
                    var (sLuP, sLuC)   = await KbGongWeiSiHuaQuery(palacesYdz, "夫妻宮", "化祿");
                    var (sQuanP,sQuanC)= await KbGongWeiSiHuaQuery(palacesYdz, "夫妻宮", "化權");
                    var (sKeP, sKeC)   = await KbGongWeiSiHuaQuery(palacesYdz, "夫妻宮", "化科");
                    var (sJiP, sJiC)   = await KbGongWeiSiHuaQuery(palacesYdz, "夫妻宮", "化忌");
                    var (hLuP, hLuC)   = await KbGongWeiSiHuaQuery(palacesYdz, "疾厄宮", "化祿");
                    var (hQuanP,hQuanC)= await KbGongWeiSiHuaQuery(palacesYdz, "疾厄宮", "化權");
                    var (hKeP, hKeC)   = await KbGongWeiSiHuaQuery(palacesYdz, "疾厄宮", "化科");
                    var (hJiP, hJiC)   = await KbGongWeiSiHuaQuery(palacesYdz, "疾厄宮", "化忌");
                    var (qQuanP,qQuanC)= await KbGongWeiSiHuaQuery(palacesYdz, "遷移宮", "化權");
                    var (qKeP, qKeC)   = await KbGongWeiSiHuaQuery(palacesYdz, "遷移宮", "化科");
                    var (qJiP, qJiC)   = await KbGongWeiSiHuaQuery(palacesYdz, "遷移宮", "化忌");
                    siHuaYdz["命宮化祿"] = (mLuP,   mLuC);   siHuaYdz["命宮化權"] = (mQuanP, mQuanC);
                    siHuaYdz["命宮化科"] = (mKeP,   mKeC);   siHuaYdz["命宮化忌"] = (mJiP,   mJiC);
                    siHuaYdz["官祿化祿"] = (oLuP,   oLuC);   siHuaYdz["官祿化權"] = (oQuanP, oQuanC);
                    siHuaYdz["官祿化科"] = (oKeP,   oKeC);   siHuaYdz["官祿化忌"] = (oJiP,   oJiC);
                    siHuaYdz["財帛化祿"] = (wLuP,   wLuC);   siHuaYdz["財帛化權"] = (wQuanP, wQuanC);
                    siHuaYdz["財帛化科"] = (wKeP,   wKeC);   siHuaYdz["財帛化忌"] = (wJiP,   wJiC);
                    siHuaYdz["夫妻化祿"] = (sLuP,   sLuC);   siHuaYdz["夫妻化權"] = (sQuanP, sQuanC);
                    siHuaYdz["夫妻化科"] = (sKeP,   sKeC);   siHuaYdz["夫妻化忌"] = (sJiP,   sJiC);
                    siHuaYdz["疾厄化祿"] = (hLuP,   hLuC);   siHuaYdz["疾厄化權"] = (hQuanP, hQuanC);
                    siHuaYdz["疾厄化科"] = (hKeP,   hKeC);   siHuaYdz["疾厄化忌"] = (hJiP,   hJiC);
                    siHuaYdz["遷移化權"] = (qQuanP, qQuanC); siHuaYdz["遷移化科"] = (qKeP,   qKeC);
                    siHuaYdz["遷移化忌"] = (qJiP,   qJiC);
                    // 先天四化（年干）
                    nianSiHuaXingYdz = await KbQuery($"SELECT COALESCE(\"ResultText\",'') AS \"Value\" FROM \"FortuneRules\" WHERE \"SourceFile\"='四化干性.docx' AND \"Title\" LIKE '{yStem}年干%' LIMIT 1");
                    siHuaLuPalaceYdz   = KbGetSiHuaPalace(yStem, "化祿", palacesYdz);
                    siHuaQuanPalaceYdz = KbGetSiHuaPalace(yStem, "化權", palacesYdz);
                    siHuaKePalaceYdz   = KbGetSiHuaPalace(yStem, "化科", palacesYdz);
                    siHuaJiPalaceYdz   = KbGetSiHuaPalace(yStem, "化忌", palacesYdz);
                    siHuaLuYdz   = await KbSiHuaQuery(yStem, "化祿", palacesYdz);
                    siHuaQuanYdz = await KbSiHuaQuery(yStem, "化權", palacesYdz);
                    siHuaKeYdz   = await KbSiHuaQuery(yStem, "化科", palacesYdz);
                    siHuaJiYdz   = await KbSiHuaQuery(yStem, "化忌", palacesYdz);
                }

                // 紫微格局偵測 + 描述查詢
                string ziweiGeJuYdz = "";
                if (hasZiwei)
                {
                    string mingBrYdz = KbGetPalaceBranch(palacesYdz, "命宮");
                    var gjList = LfDetectZiweiGeJu(mingGongStarsYdz, mingBrYdz, chartStarsYdz,
                        siHuaLuPalaceYdz, siHuaQuanPalaceYdz, siHuaKePalaceYdz, palacesYdz);
                    var gjSb = new StringBuilder();
                    foreach (var gj in gjList)
                    {
                        string desc = await KbQuery($"SELECT COALESCE(\"ResultText\",'') AS \"Value\" FROM \"FortuneRules\" WHERE \"SourceFile\"='紫微格局說明.docx' AND \"Title\"='{gj}' LIMIT 1");
                        if (!string.IsNullOrEmpty(desc)) { gjSb.AppendLine($"【{gj}】"); gjSb.AppendLine(desc); gjSb.AppendLine(); }
                    }
                    ziweiGeJuYdz = gjSb.ToString();
                }

                // 雙星組合 + 輔星入宮（5宮）
                var doubleDescsYdz = new Dictionary<string, string>();
                var minorDescsYdz  = new Dictionary<string, string>();
                if (hasZiwei)
                {
                    foreach (var kbPal in new[] { "命宮", "官祿宮", "財帛宮", "夫妻宮", "疾厄宮" })
                    {
                        doubleDescsYdz[kbPal] = await KbQueryDoubleStarInPalace(palacesYdz, kbPal);
                        minorDescsYdz[kbPal]  = await KbQueryMinorStarsInPalace(palacesYdz, kbPal);
                    }
                }

                // 十二宮星情特質（Ch.7 用）
                var allPalaceStarDescsYdz = new Dictionary<string, string>();
                if (hasZiwei)
                {
                    allPalaceStarDescsYdz["命宮"]   = starDescMingYdz;
                    allPalaceStarDescsYdz["官祿宮"] = starDescOffYdz;
                    allPalaceStarDescsYdz["財帛宮"] = starDescWltYdz;
                    allPalaceStarDescsYdz["夫妻宮"] = starDescSpsYdz;
                    allPalaceStarDescsYdz["疾厄宮"] = starDescHltYdz;
                    foreach (var pal in new[] { "兄弟宮", "子女宮", "遷移宮", "奴僕宮", "田宅宮", "福德宮", "父母宮" })
                        allPalaceStarDescsYdz[pal] = await KbQueryStarInPalace(palacesYdz, pal);
                }

                string reportText = LfBuildYudongziReportV2(
                    yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch,
                    yStemSS, mStemSS, hStemSS, yBranchSS, mBranchSS, dBranchSS, hBranchSS,
                    yNaYin, mNaYin, dNaYin, hNaYin,
                    dmElem, wuXing, bodyPct, bodyLabel, season, seaLabel,
                    pattern, yongShenElem, fuYiElem, yongReason, jiShenElem,
                    scored, gender, birthYear, user.BirthMonth, user.BirthDay, user.BirthHour, user.BirthMinute, lunarMonthDocx,
                    hasZiwei, palacesYdz, mingGongStarsYdz, mingZhuYdz, shenZhuYdz, wuXingJuTextYdz,
                    ziweiMingYdz, starDescMingYdz, ziweiFullContentYdz, chartStarsYdz,
                    ziweiOffYdz, offStarsYdz, ziweiWltYdz, wltStarsYdz,
                    ziweiSpsYdz, spsStarsYdz, ziweiHltYdz, hltStarsYdz,
                    siHuaYdz,
                    nianSiHuaXingYdz, yStem,
                    siHuaLuPalaceYdz, siHuaLuYdz,
                    siHuaQuanPalaceYdz, siHuaQuanYdz,
                    siHuaKePalaceYdz, siHuaKeYdz,
                    siHuaJiPalaceYdz, siHuaJiYdz,
                    dayPillarKb, zhongyuanRules, ziweiGeJuYdz,
                    doubleDescsYdz, minorDescsYdz, allPalaceStarDescsYdz,
                    starDescOffYdz, starDescWltYdz, starDescSpsYdz, starDescHltYdz,
                    ziweiParStarYdz, ziweiParYdz, ziweiCldStarYdz, ziweiCldYdz,
                    userName: docxUserName,
                    calDb: _calendarDb);

                // === 建立 DOCX ===
                string wwwroot = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot");
                string coverPath = Path.Combine(wwwroot, "images", "cover_page.jpg");
                string sealPath  = Path.Combine(wwwroot, "images", "玉洞子印.png");
                byte[] coverBytes = System.IO.File.Exists(coverPath) ? await System.IO.File.ReadAllBytesAsync(coverPath) : Array.Empty<byte>();
                byte[] sealBytes  = System.IO.File.Exists(sealPath)  ? await System.IO.File.ReadAllBytesAsync(sealPath)  : Array.Empty<byte>();
                byte[] chartImgBytes = string.IsNullOrEmpty(request?.ChartImageBase64)
                    ? Array.Empty<byte>()
                    : Convert.FromBase64String(request.ChartImageBase64);

                // 九星氣學加成（純 KB，append 至 reportText）
                string docxNsSection = await NsBuildBirthSection(
                    user.BirthYear ?? birthYear,
                    user.BirthMonth ?? 1,
                    user.BirthDay ?? 1,
                    user.BirthHour ?? 0,
                    user.BirthGender ?? gender);
                if (!string.IsNullOrEmpty(docxNsSection)) reportText += docxNsSection;

                string personName = !string.IsNullOrEmpty(request.PersonName) ? request.PersonName : (user.Name ?? "命主");
                byte[] docxBytes = LfBuildYudongziDocxBytes(reportText, coverBytes, chartImgBytes, sealBytes, personName, "玉 洞 子 傳 家 寶 典");

                string fileName = $"{personName}_玉洞子傳家寶典.docx";
                return File(docxBytes,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    fileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "玉洞子命書DOCX失敗 User={User}", identity);
                return StatusCode(500, new { error = "DOCX生成失敗", details = ex.Message });
            }
        }

        private static byte[] LfBuildYudongziDocxBytes(string reportText, byte[] coverBytes, byte[] chartImgBytes, byte[] sealBytes, string personName, string bookTitle = "玉 洞 子 傳 家 寶 典", string? skipTitle = null)
        {
            using var ms = new MemoryStream();
            using var doc = new NPOI.XWPF.UserModel.XWPFDocument();

            void AddPara(string text, int fontSize, bool bold, string colorHex, NPOI.XWPF.UserModel.ParagraphAlignment align)
            {
                var p = doc.CreateParagraph();
                p.Alignment = align;
                if (bold) p.SpacingBefore = 80;
                var r = p.CreateRun();
                r.SetFontFamily("標楷體", NPOI.XWPF.UserModel.FontCharRange.None);
                r.FontSize = fontSize;
                r.IsBold = bold;
                r.SetColor(colorHex);
                r.SetText(text ?? "");
            }

            void AddImage(byte[] imgBytes, int pictureType, int widthCm, int heightCm, bool leftAlign = false)
            {
                if (imgBytes == null || imgBytes.Length == 0) return;
                var p = doc.CreateParagraph();
                p.Alignment = leftAlign ? NPOI.XWPF.UserModel.ParagraphAlignment.LEFT : NPOI.XWPF.UserModel.ParagraphAlignment.CENTER;
                var r = p.CreateRun();
                using var imgStream = new MemoryStream(imgBytes);
                int wEmu = widthCm * 360000;
                int hEmu = heightCm * 360000;
                r.AddPicture(imgStream, pictureType, "img", wEmu, hEmu);
            }

            void AddPageBreak()
            {
                var p = doc.CreateParagraph();
                var r = p.CreateRun();
                r.AddBreak(NPOI.XWPF.UserModel.BreakType.PAGE);
            }

            // 用 IsPageBreak 段落屬性換頁：段落從新頁開始，不在前頁留空白段落
            void AddParaWithPageBreak(string text, int fontSize, bool bold, string colorHex, NPOI.XWPF.UserModel.ParagraphAlignment align)
            {
                var p = doc.CreateParagraph();
                p.Alignment = align;
                if (bold) p.SpacingBefore = 80;
                p.IsPageBreak = true; // 段落從新頁起始，根治空白頁問題
                var r = p.CreateRun();
                r.SetFontFamily("標楷體", NPOI.XWPF.UserModel.FontCharRange.None);
                r.FontSize = fontSize;
                r.IsBold = bold;
                r.SetColor(colorHex);
                r.SetText(text ?? "");
            }

            // 章節標題（帶換頁 + 大綱層級 0，供 Word TOC \u 識別）
            void AddParaWithPageBreakH1(string text, int fontSize, bool bold, string colorHex, NPOI.XWPF.UserModel.ParagraphAlignment align)
            {
                var p = doc.CreateParagraph();
                p.Alignment = align;
                if (bold) p.SpacingBefore = 80;
                p.IsPageBreak = true;
                // 設定 outline level 0（= Heading 1）讓 Word TOC \u 識別
                var pPr = p.GetCTP().pPr ?? p.GetCTP().AddNewPPr();
                var ol = new NPOI.OpenXmlFormats.Wordprocessing.CT_DecimalNumber();
                ol.val = "0";
                pPr.outlineLvl = ol;
                var r = p.CreateRun();
                r.SetFontFamily("標楷體", NPOI.XWPF.UserModel.FontCharRange.None);
                r.FontSize = fontSize;
                r.IsBold = bold;
                r.SetColor(colorHex);
                r.SetText(text ?? "");
            }

            // 插入 Word TOC 欄位（\u=大綱層級, \h=超連結, \z=隱藏Web版頁碼）
            void InsertWordTocField()
            {
                var para = doc.CreateParagraph();
                para.Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.LEFT;
                var ctP = para.GetCTP();
                var r1 = ctP.AddNewR();
                r1.AddNewFldChar().fldCharType = NPOI.OpenXmlFormats.Wordprocessing.ST_FldCharType.begin;
                var r2 = ctP.AddNewR();
                var instr = r2.AddNewInstrText();
                instr.Value = " TOC \\u \\h \\z ";
                instr.space = "preserve";
                var r3 = ctP.AddNewR();
                r3.AddNewFldChar().fldCharType = NPOI.OpenXmlFormats.Wordprocessing.ST_FldCharType.separate;
                var r4 = ctP.AddNewR();
                r4.AddNewT().Value = "（請在 Word 中按 F9 更新目錄）";
                var r5 = ctP.AddNewR();
                r5.AddNewFldChar().fldCharType = NPOI.OpenXmlFormats.Wordprocessing.ST_FldCharType.end;
            }

            // === 封面（三欄：左聯 | 中央 | 右聯）===
            string coverImgDir = Path.Combine(AppContext.BaseDirectory, "wwwroot", "images");
            byte[] scrollLeftBytes  = System.IO.File.Exists(Path.Combine(coverImgDir, "scroll_left.jpg"))  ? System.IO.File.ReadAllBytes(Path.Combine(coverImgDir, "scroll_left.jpg"))  : Array.Empty<byte>();
            byte[] scrollRightBytes = System.IO.File.Exists(Path.Combine(coverImgDir, "scroll_right.jpg")) ? System.IO.File.ReadAllBytes(Path.Combine(coverImgDir, "scroll_right.jpg")) : Array.Empty<byte>();

            if (scrollLeftBytes.Length > 0 && scrollRightBytes.Length > 0)
            {
                var coverTbl = doc.CreateTable(1, 3);
                // 移除所有邊框
                var ctTblPr = coverTbl.GetCTTbl().tblPr ?? coverTbl.GetCTTbl().AddNewTblPr();
                var ctBrd = ctTblPr.tblBorders ?? ctTblPr.AddNewTblBorders();
                void NoBorder(NPOI.OpenXmlFormats.Wordprocessing.CT_Border b)
                { b.val = NPOI.OpenXmlFormats.Wordprocessing.ST_Border.none; b.sz = 0; b.space = 0; b.color = "auto"; }
                NoBorder(ctBrd.AddNewTop()); NoBorder(ctBrd.AddNewBottom());
                NoBorder(ctBrd.AddNewLeft()); NoBorder(ctBrd.AddNewRight());
                NoBorder(ctBrd.AddNewInsideH()); NoBorder(ctBrd.AddNewInsideV());

                // 紅底色 helper
                void SetCellRedBg(NPOI.XWPF.UserModel.XWPFTableCell cell)
                {
                    var tcPr = cell.GetCTTc().tcPr ?? cell.GetCTTc().AddNewTcPr();
                    var shd  = tcPr.shd ?? tcPr.AddNewShd();
                    shd.val   = NPOI.OpenXmlFormats.Wordprocessing.ST_Shd.clear;
                    shd.color = "auto";
                    shd.fill  = "8B0000"; // 深紅底
                }

                void SetCellWidth(NPOI.XWPF.UserModel.XWPFTableCell cell, int dxa)
                {
                    var tcPr = cell.GetCTTc().tcPr ?? cell.GetCTTc().AddNewTcPr();
                    var tcW  = tcPr.tcW ?? tcPr.AddNewTcW();
                    tcW.type = NPOI.OpenXmlFormats.Wordprocessing.ST_TblWidth.dxa;
                    tcW.w    = dxa.ToString();
                    var va   = tcPr.vAlign ?? tcPr.AddNewVAlign();
                    va.val   = NPOI.OpenXmlFormats.Wordprocessing.ST_VerticalJc.center;
                }

                void AddCellPara(NPOI.XWPF.UserModel.XWPFTableCell cell, string text, int fs, bool bold, string color)
                {
                    var p = cell.AddParagraph();
                    p.Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.CENTER;
                    var r = p.CreateRun();
                    r.SetFontFamily("標楷體", NPOI.XWPF.UserModel.FontCharRange.None);
                    r.FontSize = fs; r.IsBold = bold; r.SetColor(color); r.SetText(text);
                }

                var coverRow = coverTbl.GetRow(0);

                // 左聯（洞合乾坤養道丹）
                var lcell = coverRow.GetCell(0);
                SetCellWidth(lcell, 1985); // 3.5cm
                SetCellRedBg(lcell);
                var lp = lcell.Paragraphs.Count > 0 ? lcell.Paragraphs[0] : lcell.AddParagraph();
                lp.Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.LEFT;
                lp.IndentationLeft = 567; // 1cm 向中靠攏
                using var lstream = new MemoryStream(scrollLeftBytes);
                lp.CreateRun().AddPicture(lstream, (int)NPOI.XWPF.UserModel.PictureType.JPEG, "scroll_left", (int)(3.2 * 360000), (int)(12.0 * 360000));

                // 中央文字（金字）
                var ccell = coverRow.GetCell(1);
                SetCellWidth(ccell, 6236); // 11cm
                SetCellRedBg(ccell);
                var cp0 = ccell.Paragraphs.Count > 0 ? ccell.Paragraphs[0] : ccell.AddParagraph();
                cp0.Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.CENTER;
                cp0.CreateRun().SetText(""); // 首行佔位
                // 橫排：玉虛洞天（金色）
                AddCellPara(ccell, "", 14, false, "D4AF37");
                AddCellPara(ccell, "玉 虛 洞 天", 20, true, "D4AF37");
                AddCellPara(ccell, "", 10, false, "D4AF37");
                // 直排：命主名（逐字分行 36pt 金字）
                foreach (char c in personName)
                    AddCellPara(ccell, c.ToString(), 36, true, "FFD700");
                AddCellPara(ccell, "", 10, false, "D4AF37");
                // 直排：親鑑（逐字分行 36pt 淡金）
                AddCellPara(ccell, "親", 36, false, "D4AF37");
                AddCellPara(ccell, "鑑", 36, false, "D4AF37");

                // 右聯（玉懷天地積德心）
                var rcell = coverRow.GetCell(2);
                SetCellWidth(rcell, 1985); // 3.5cm
                SetCellRedBg(rcell);
                var rp = rcell.Paragraphs.Count > 0 ? rcell.Paragraphs[0] : rcell.AddParagraph();
                rp.Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.RIGHT;
                rp.IndentationRight = 567; // 1cm 向中靠攏
                using var rstream = new MemoryStream(scrollRightBytes);
                rp.CreateRun().AddPicture(rstream, (int)NPOI.XWPF.UserModel.PictureType.JPEG, "scroll_right", (int)(3.2 * 360000), (int)(12.0 * 360000));
            }
            else
            {
                // 備用封面（無對聯圖片）
                if (coverBytes.Length > 0)
                    AddImage(coverBytes, (int)NPOI.XWPF.UserModel.PictureType.JPEG, 16, 8);
                AddPara(bookTitle, 36, true, "8B0000", NPOI.XWPF.UserModel.ParagraphAlignment.CENTER);
                AddPara("親鑑", 16, false, "8B4513", NPOI.XWPF.UserModel.ParagraphAlignment.CENTER);
                AddPara($"命主：{personName}", 20, true, "000000", NPOI.XWPF.UserModel.ParagraphAlignment.CENTER);
                AddPara("  時辰恐有錯  陰騭最難憑", 13, false, "CC0000", NPOI.XWPF.UserModel.ParagraphAlignment.CENTER);
                AddPara("  萬般皆是命  半點不求人", 13, false, "CC0000", NPOI.XWPF.UserModel.ParagraphAlignment.CENTER);
            }
            AddPageBreak();

            // === 第二頁：原封面（保留）===
            if (coverBytes.Length > 0)
                AddImage(coverBytes, (int)NPOI.XWPF.UserModel.PictureType.JPEG, 16, 8);
            AddPara(bookTitle, 36, true, "8B0000", NPOI.XWPF.UserModel.ParagraphAlignment.CENTER);
            AddPara("親鑑", 16, false, "8B4513", NPOI.XWPF.UserModel.ParagraphAlignment.CENTER);
            AddPara($"命主：{personName}", 20, true, "000000", NPOI.XWPF.UserModel.ParagraphAlignment.CENTER);
            AddPara(" ", 12, false, "000000", NPOI.XWPF.UserModel.ParagraphAlignment.LEFT);
            AddPara("  時辰恐有錯  陰騭最難憑", 13, false, "CC0000", NPOI.XWPF.UserModel.ParagraphAlignment.CENTER);
            AddPara("  萬般皆是命  半點不求人", 13, false, "CC0000", NPOI.XWPF.UserModel.ParagraphAlignment.CENTER);
            AddPageBreak();

            // === 先天元神圖 ===
            if (chartImgBytes.Length > 0)
            {
                AddPara("【先天元神圖】", 18, true, "8B0000", NPOI.XWPF.UserModel.ParagraphAlignment.CENTER);
                AddPara(" ", 8, false, "000000", NPOI.XWPF.UserModel.ParagraphAlignment.LEFT);
                AddImage(chartImgBytes, (int)NPOI.XWPF.UserModel.PictureType.PNG, 15, 15, leftAlign: true);
                AddPageBreak();
            }

            // === 16 章報告 ===
            var pipeBuffer = new List<string>();
            int curTableFontSize = 10; // 預設表格字型，第二章改為18

            void FlushPipeTable()
            {
                if (pipeBuffer.Count == 0) return;
                var dataRows = new List<List<string>>();
                foreach (var pl in pipeBuffer)
                {
                    // skip separator rows (|---|---|)
                    if (pl.Replace("|","").Replace("-","").Replace(":","").Trim() == "") continue;
                    var parts = pl.Split('|').Skip(1).ToList();
                    if (parts.Count > 0 && parts.Last().Trim() == "") parts.RemoveAt(parts.Count - 1);
                    var cells = parts.Select(s => s.Trim()).ToList();
                    if (cells.Count > 0) dataRows.Add(cells);
                }
                pipeBuffer.Clear();
                if (dataRows.Count == 0) return;

                int colCount = dataRows.Max(r => r.Count);
                var tbl = doc.CreateTable(dataRows.Count, colCount);
                for (int ri = 0; ri < dataRows.Count; ri++)
                {
                    var trow = tbl.GetRow(ri);
                    for (int ci = 0; ci < dataRows[ri].Count && ci < colCount; ci++)
                    {
                        var tcell = trow.GetCell(ci);
                        var para = tcell.Paragraphs.Count > 0 ? tcell.Paragraphs[0] : tcell.AddParagraph();
                        para.Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.CENTER;
                        var run = para.CreateRun();
                        run.SetFontFamily("標楷體", NPOI.XWPF.UserModel.FontCharRange.None);
                        run.FontSize = curTableFontSize;
                        if (ri == 0) run.IsBold = true;
                        run.SetText(dataRows[ri][ci]);
                    }
                }
                doc.CreateParagraph(); // spacing after table
            }

            string effectiveSkip = skipTitle ?? bookTitle;
            bool ShouldSkipReportLine(string line)
            {
                if (line.Contains(effectiveSkip)) return true;
                if (line.TrimEnd() == "  時辰恐有錯  陰騭最難憑") return true;
                if (line.TrimEnd() == "  萬般皆是命  半點不求人") return true;
                if (line.Contains("命理大師：玉洞子")) return true;
                if (line.StartsWith("性別：") && line.Contains("虛齡")) return true;
                if (line.StartsWith("四柱：") && line.Contains(" ")) return true;
                return false;
            }

            int kbSectionCount = 0; // 計算 === === 章節數，用於換頁
            bool inTocSection = false; // 目前在純文字目錄區間，略過這些行

            // 過濾：移除緊接在章節標題前的空白行，防止 DOCX 出現空白頁
            // 向前看時跳過所有最終會被 DOCX 忽略的行（空白行、=====、-----、頁尾簽名）
            var rawLines = reportText.Split('\n');
            var filteredLines = new List<string>(rawLines.Length);
            bool IsSkippableLine(string s) =>
                string.IsNullOrWhiteSpace(s) ||
                s.StartsWith("=====") ||
                s.StartsWith("-----") ||
                ShouldSkipReportLine(s);
            for (int li = 0; li < rawLines.Length; li++)
            {
                string lt = rawLines[li].TrimEnd();
                if (string.IsNullOrWhiteSpace(lt))
                {
                    int nxt = li + 1;
                    while (nxt < rawLines.Length && IsSkippableLine(rawLines[nxt].TrimEnd())) nxt++;
                    if (nxt < rawLines.Length)
                    {
                        string nl = rawLines[nxt].TrimEnd();
                        if (nl.StartsWith("【第") && nl.EndsWith("】")) continue;
                    }
                }
                filteredLines.Add(lt);
            }

            foreach (var line in filteredLines)
            {

                if (ShouldSkipReportLine(line)) continue;

                // 偵測「人  生  指  南」→ 插入 Word TOC 欄位，略過純文字目錄項目
                if (line.TrimEnd().Contains("人  生  指  南"))
                {
                    AddPara("人  生  指  南", 18, true, "8B0000", NPOI.XWPF.UserModel.ParagraphAlignment.CENTER);
                    InsertWordTocField();
                    inTocSection = true;
                    continue;
                }
                if (inTocSection)
                {
                    if (line.StartsWith("【第") && line.EndsWith("】"))
                        inTocSection = false; // 遇到第一章，結束目錄區間，繼續往下處理
                    else
                        continue; // 略過純文字目錄項目
                }

                // Buffer pipe-table lines
                if (line.TrimStart().StartsWith("|"))
                {
                    pipeBuffer.Add(line);
                    continue;
                }

                // Non-pipe line: flush any buffered table first
                FlushPipeTable();

                if (line == "【第二章：先天八字依古制定】" || line.StartsWith("【第三章：日柱深度論斷") || line == "【第三章：深度分析】" ||
                    line == "【第四章：命格判定】" || line == "【第五章：用神喜忌】" ||
                    line == "【第六章：紫微星格】" || line == "【第七章：宮星化象（十二宮）】")
                {
                    // 第二章表格用18pt，其他章節恢復10pt
                    curTableFontSize = (line == "【第二章：先天八字依古制定】") ? 18 : 10;
                    AddParaWithPageBreakH1(line, 16, true, "8B0000", NPOI.XWPF.UserModel.ParagraphAlignment.LEFT);
                }
                else if (line.StartsWith("【第") && line.EndsWith("】"))
                {
                    curTableFontSize = 10; // 非第二章的其他章節，恢復預設
                    // 所有書第1章以後自動換頁（加了目錄後，第1章也需換頁）
                    var chM = System.Text.RegularExpressions.Regex.Match(line, @"【第([一二三四五六七八九十\d]+)章");
                    if (chM.Success)
                    {
                        string chStr = chM.Groups[1].Value;
                        int chN = int.TryParse(chStr, out var cn) ? cn : LfChineseNumToInt(chStr);
                        if (chN >= 1)
                            AddParaWithPageBreakH1(line, 16, true, "8B0000", NPOI.XWPF.UserModel.ParagraphAlignment.LEFT);
                        else
                            AddPara(line, 16, true, "8B0000", NPOI.XWPF.UserModel.ParagraphAlignment.LEFT);
                        continue;
                    }
                    AddPara(line, 16, true, "8B0000", NPOI.XWPF.UserModel.ParagraphAlignment.LEFT);
                }
                else if (System.Text.RegularExpressions.Regex.IsMatch(line.Trim(), @"^={2,}\s+.+\s+={2,}$"))
                {
                    // === 一、XXX === 格式：KB 報告章節標題
                    var secTitle = System.Text.RegularExpressions.Regex.Replace(line.Trim(), @"^=+\s*|\s*=+$", "").Trim();
                    if (kbSectionCount++ > 0) AddPageBreak();
                    AddPara(secTitle, 14, true, "8B4513", NPOI.XWPF.UserModel.ParagraphAlignment.LEFT);
                }
                else if (System.Text.RegularExpressions.Regex.IsMatch(line.Trim(), @"^-{2,}\s+.+\s+-{2,}$"))
                {
                    // --- XXX --- 格式：KB 報告子標題
                    var subTitle = System.Text.RegularExpressions.Regex.Replace(line.Trim(), @"^-+\s*|\s*-+$", "").Trim();
                    AddPara(subTitle, 12, true, "5C3317", NPOI.XWPF.UserModel.ParagraphAlignment.LEFT);
                }
                else if (line.StartsWith("=====") || line.StartsWith("-----"))
                {
                    // skip pure divider lines from report header block
                }
                else if (string.IsNullOrWhiteSpace(line))
                {
                    AddPara(" ", 8, false, "000000", NPOI.XWPF.UserModel.ParagraphAlignment.LEFT);
                }
                else if (line.StartsWith("【") && line.Contains("："))
                {
                    // 若【label】在行中間（非結尾），為內嵌標籤格式，套用內文字型
                    int closeIdx = line.IndexOf("】");
                    bool isInlineLabel = closeIdx >= 0 && closeIdx < line.Length - 1;
                    if (isInlineLabel)
                        AddPara(line, 11, false, "222222", NPOI.XWPF.UserModel.ParagraphAlignment.LEFT);
                    else
                        AddPara(line, 13, true, "5C3317", NPOI.XWPF.UserModel.ParagraphAlignment.LEFT);
                }
                else
                {
                    AddPara(line, 11, false, "222222", NPOI.XWPF.UserModel.ParagraphAlignment.LEFT);
                }
            }
            FlushPipeTable(); // flush any trailing table

            // === 玉洞子印（跳頁 + 共勉語 + 印章）===
            AddPageBreak();
            AddPara(" ", 12, false, "000000", NPOI.XWPF.UserModel.ParagraphAlignment.LEFT);
            AddPara("算命的真蹄", 22, true, "333333", NPOI.XWPF.UserModel.ParagraphAlignment.LEFT);
            AddPara(" ", 12, false, "000000", NPOI.XWPF.UserModel.ParagraphAlignment.LEFT);
            AddPara("在求內因的充實及外緣的爭取", 20, false, "333333", NPOI.XWPF.UserModel.ParagraphAlignment.LEFT);
            AddPara(" ", 12, false, "000000", NPOI.XWPF.UserModel.ParagraphAlignment.LEFT);
            AddPara("而不在準與不準　六親隨緣，人生隨運", 20, false, "333333", NPOI.XWPF.UserModel.ParagraphAlignment.LEFT);
            AddPara(" ", 12, false, "000000", NPOI.XWPF.UserModel.ParagraphAlignment.LEFT);
            AddPara(" ", 12, false, "000000", NPOI.XWPF.UserModel.ParagraphAlignment.LEFT);
            AddPara("　　　　　　共 勉 之", 20, false, "333333", NPOI.XWPF.UserModel.ParagraphAlignment.LEFT);
            AddPara(" ", 12, false, "000000", NPOI.XWPF.UserModel.ParagraphAlignment.LEFT);
            if (sealBytes.Length > 0)
            {
                var p = doc.CreateParagraph();
                p.Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.CENTER;
                var r = p.CreateRun();
                using var sealStream = new MemoryStream(sealBytes);
                r.AddPicture(sealStream, (int)NPOI.XWPF.UserModel.PictureType.PNG, "seal", 3600000, 3600000);
            }
            else
            {
                AddPara("　　　　[ 玉 洞 子 印 ]", 18, true, "CC0000", NPOI.XWPF.UserModel.ParagraphAlignment.CENTER);
            }
            AddPara(" 敬 批", 16, true, "CC0000", NPOI.XWPF.UserModel.ParagraphAlignment.CENTER);

            doc.Write(ms);
            return ms.ToArray();
        }

        // === 通用命書 DOCX 匯出 ===

        public class GenericDocxRequest
        {
            public string ReportText  { get; set; } = "";
            public string PersonName  { get; set; } = "";
            public string BookTitle   { get; set; } = "";   // 封面顯示標題
            public string? SkipTitle  { get; set; }         // report body 裡要略過的標題行（預設同 BookTitle）
            public string? ChartImageBase64 { get; set; }
        }

        [HttpPost("export-generic-docx")]
        [Authorize]
        public async Task<IActionResult> ExportGenericDocx([FromBody] GenericDocxRequest request)
        {
            try
            {
                string wwwroot   = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot");
                string coverPath = Path.Combine(wwwroot, "images", "cover_page.jpg");
                string sealPath  = Path.Combine(wwwroot, "images", "玉洞子印.png");
                byte[] coverBytes = System.IO.File.Exists(coverPath) ? await System.IO.File.ReadAllBytesAsync(coverPath) : Array.Empty<byte>();
                byte[] sealBytes  = System.IO.File.Exists(sealPath)  ? await System.IO.File.ReadAllBytesAsync(sealPath)  : Array.Empty<byte>();
                byte[] chartImgBytes = string.IsNullOrEmpty(request.ChartImageBase64)
                    ? Array.Empty<byte>()
                    : Convert.FromBase64String(request.ChartImageBase64);

                string personName = !string.IsNullOrEmpty(request.PersonName) ? request.PersonName : "命主";
                string bookTitle  = !string.IsNullOrEmpty(request.BookTitle)  ? request.BookTitle  : "命書";

                string skipTitle = !string.IsNullOrEmpty(request.SkipTitle) ? request.SkipTitle : bookTitle;
                byte[] docxBytes = LfBuildYudongziDocxBytes(request.ReportText, coverBytes, chartImgBytes, sealBytes, personName, bookTitle, skipTitle);

                string safeTitle = bookTitle.Replace(" ", "");
                string fileName  = $"{personName}_{safeTitle}.docx";
                return File(docxBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", fileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "GenericDocx失敗");
                return StatusCode(500, new { error = "DOCX生成失敗", details = ex.Message });
            }
        }

        // === Dy (大運命書) Endpoint ===

        [HttpGet("analyze-daiyun")]
        [Authorize]
        public async Task<IActionResult> GetDaiyunAnalysis([FromQuery] int years = 5)
        {
            var identity = User.FindFirstValue(ClaimTypes.Email)
                         ?? User.FindFirstValue(ClaimTypes.Name)
                         ?? User.FindFirst("unique_name")?.Value;
            if (string.IsNullOrEmpty(identity))
                return Unauthorized(new { error = "請重新登入" });

            var user = await _context.Users.FirstOrDefaultAsync(u => u.UserName == identity || u.Email == identity);
            if (user == null) return BadRequest(new { error = "找不到用戶" });

            var userChart = await _context.UserCharts.FirstOrDefaultAsync(c => c.UserId == user.Id);
            if (userChart == null || string.IsNullOrEmpty(userChart.ChartJson))
                return BadRequest(new { error = "no_chart" });

            bool dyIsAdmin = string.Equals(user.Email, _config["Admin:Email"], StringComparison.OrdinalIgnoreCase);
            int dySubId = -1;
            if (!dyIsAdmin)
            {
                var (dyOk, dyErr, dySubIdVal) = await CheckSubscriptionQuota(user.Id, "BOOK_DAIYUN");
                if (!dyOk) return BadRequest(new { error = dyErr });
                dySubId = dySubIdVal;
            }

            try
            {
                var root = JsonDocument.Parse(userChart.ChartJson).RootElement;
                if (!root.TryGetProperty("bazi", out var bazi) && !root.TryGetProperty("baziInfo", out bazi))
                    return BadRequest(new { error = "命盤資料格式錯誤" });

                var yearP  = LfGetPillar(bazi, "yearPillar");
                var monthP = LfGetPillar(bazi, "monthPillar");
                var dayP   = LfGetPillar(bazi, "dayPillar");
                var timeP  = LfGetPillar(bazi, "timePillar");

                string yStem = LfPillarStem(yearP);   string yBranch = LfPillarBranch(yearP);
                string mStem = LfPillarStem(monthP);  string mBranch = LfPillarBranch(monthP);
                string dStem = LfPillarStem(dayP);    string dBranch = LfPillarBranch(dayP);
                string hStem = LfPillarStem(timeP);   string hBranch = LfPillarBranch(timeP);

                string yStemSS   = LfPillarStemSS(yearP);
                string mStemSS   = LfPillarStemSS(monthP);
                string hStemSS   = LfPillarStemSS(timeP);
                string yBranchSS = LfPillarBranchMainSS(yearP);
                string mBranchSS = LfPillarBranchMainSS(monthP);
                string dBranchSS = LfPillarBranchMainSS(dayP);
                string hBranchSS = LfPillarBranchMainSS(timeP);

                int birthYear = user.BirthYear ?? (DateTime.Today.Year - 30);
                int gender    = user.BirthGender ?? 1;
                string dmElem = KbStemToElement(dStem);
                var branches  = new[] { yBranch, mBranch, dBranch, hBranch };
                var wuXing    = LfCalcWuXingMatrix(yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch);
                double bodyPct   = LfGetBodyStrengthPct(dmElem, wuXing);
                string bodyLabel = LfGetBodyStrengthLabel(bodyPct);
                string season    = LfGetSeason(mBranch);
                string seaLabel  = LfGetSeasonLabel(mBranch);

                var (pattern, yongShenElem, fuYiElem, yongReason, tiaoHouElem) = LfDetectGeJuAndYongShen(
                    yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch,
                    dmElem, wuXing, bodyPct, season);
                string jiShenElem = LfGetJiShenElem(yongShenElem, dmElem, bodyPct, pattern);

                var luckCycles = LfExtractLuckCycles(root);
                bool hasZiwei  = root.TryGetProperty("palaces", out var palaces) && palaces.ValueKind == JsonValueKind.Array;

                int startYear = DateTime.Today.Year;
                int endYear   = years == 0 ? Math.Max(startYear + 1, birthYear + 80) : startYear + years - 1;

                var chartStems3 = new[] { yStem, mStem, dStem, hStem };
                var annualDetails = new List<(int year, int age, string flStem, string flBranch,
                    string daiyunStem, string daiyunBranch,
                    int baziScore, int ziweiScore, string crossClass,
                    string flStemSS, string flBranchSS)>();
                var annualForecasts = new List<object>();

                for (int yr = startYear; yr <= endYear; yr++)
                {
                    int age = yr - birthYear;
                    var (flStem, flBranch) = DyGetYearStemBranch(yr);
                    var curLuck = luckCycles.FirstOrDefault(lc => age >= lc.startAge && age < lc.endAge);
                    string daiyunStem   = curLuck.stem   ?? "";
                    string daiyunBranch = curLuck.branch ?? "";

                    int baziScore  = DyCalcFlowYearBaziScore(flStem, flBranch, pattern, yongShenElem, fuYiElem, jiShenElem,
                        dmElem, bodyPct > 50, tiaoHouElem, season, branches, chartStems3);
                    int ziweiScore = hasZiwei ? DyCalcZiweiScore(flStem, palaces, daiyunStem, age) : 50;
                    string crossClass = DyCrossClass(baziScore, ziweiScore);

                    string flBrMs = LfBranchHiddenRatio.TryGetValue(flBranch, out var fbh) && fbh.Count > 0 ? fbh[0].stem : "";
                    string flStemSS   = LfStemShiShen(flStem, dStem);
                    string flBranchSS = !string.IsNullOrEmpty(flBrMs) ? LfStemShiShen(flBrMs, dStem) : "";

                    annualDetails.Add((yr, age, flStem, flBranch, daiyunStem, daiyunBranch,
                        baziScore, ziweiScore, crossClass, flStemSS, flBranchSS));
                    annualForecasts.Add(new {
                        year = yr, age, stemBranch = flStem + flBranch,
                        daiyunStem, daiyunBranch, baziScore, ziweiScore, crossClass,
                        summary = DyYearSummary(crossClass, flStemSS, flBranchSS, baziScore, ziweiScore),
                        detail  = DyCrossDesc(crossClass, flStemSS, flBranchSS, baziScore, ziweiScore)
                    });
                }

                // 預取四化入宮描述（先天四化.docx）for all unique stems
                // 流年四化用流年天干；大限四化用大限宮位的宮干（非八字大運天干）
                // 一個八字大運可能橫跨多個紫微大限（年齡邊界不同），需 SelectMany 收集全部宮干
                var siHuaDescMap = new Dictionary<string, Dictionary<string, (string palace, string desc)>>();
                if (hasZiwei)
                {
                    var decadePalaceStems = luckCycles.SelectMany(lc =>
                        DyGetOverlappingDecadePalaces(palaces, lc.startAge, lc.endAge)
                            .Select(o => o.palaceStem));
                    var uniqueStems = annualDetails.Select(a => a.flStem)
                        .Concat(decadePalaceStems)
                        .Where(s => !string.IsNullOrEmpty(s) && YearStemSiHuaMap.ContainsKey(s))
                        .Distinct().ToList();
                    string[] siHuaTypes = { "化祿", "化權", "化科", "化忌" };
                    foreach (var stem in uniqueStems)
                    {
                        var stemMap = new Dictionary<string, (string palace, string desc)>();
                        foreach (var sh in siHuaTypes)
                        {
                            string pal  = KbGetSiHuaPalace(stem, sh, palaces);
                            string desc = string.IsNullOrEmpty(pal) ? "" : await KbSiHuaQuery(stem, sh, palaces);
                            // 截斷過長描述，取首句
                            if (desc.Length > 60)
                            {
                                int dot = desc.IndexOfAny(new[] { '。', '，', '\n' });
                                desc = dot > 0 && dot < 80 ? desc[..(dot + 1)] : desc[..60] + "...";
                            }
                            stemMap[sh] = (pal, desc);
                        }
                        siHuaDescMap[stem] = stemMap;
                    }
                }

                // baziTable（前端命盤結構卡片）
                var baziTable = new {
                    pillars = new[] {
                        new { label = "年", stem = yStem, branch = yBranch, stemSS = yStemSS,
                              naYin = LfPillarNaYin(yearP), hiddenPairs = LfPillarHiddenPairs(yearP) },
                        new { label = "月", stem = mStem, branch = mBranch, stemSS = mStemSS,
                              naYin = LfPillarNaYin(monthP), hiddenPairs = LfPillarHiddenPairs(monthP) },
                        new { label = "日", stem = dStem, branch = dBranch, stemSS = "元神",
                              naYin = LfPillarNaYin(dayP), hiddenPairs = LfPillarHiddenPairs(dayP) },
                        new { label = "時", stem = hStem, branch = hBranch, stemSS = hStemSS,
                              naYin = LfPillarNaYin(timeP), hiddenPairs = LfPillarHiddenPairs(timeP) },
                    }
                };

                // luckCycles（前端大運走勢圖）
                var scoredCycles = luckCycles.Select(lc => {
                    int sc = LfCalcLuckScore(lc.stem, lc.branch, pattern, yongShenElem, fuYiElem, jiShenElem,
                        dmElem, bodyPct > 50, tiaoHouElem, season, branches, chartStems3, dStem);
                    return new { lc.stem, lc.branch, lc.liuShen, lc.startAge, lc.endAge, score = sc, level = LfLuckLevel(sc) };
                }).ToList();

                // 預取紫微完整命盤內容（大限宮位主星格局用）
                string ziweiFullContent = "";
                var chartStars = new HashSet<string>();
                // 大限命宮 KB：以每個大限宮位地支為 palace_position 查詢，取「命宮」段落作為該大限格局
                var decadeKbMap = new Dictionary<string, string>();
                if (hasZiwei)
                {
                    string ziweiPos = KbGetZiweiPosition(palaces);
                    ziweiFullContent = await KbZiweiFullQuery(palaces, ziweiPos);
                    chartStars = KbGetAllChartStars(palaces);
                    // 收集所有大限宮位的唯一地支並預查 KB
                    var uniqueDecadeBranches = new HashSet<string>();
                    foreach (var lc in luckCycles)
                    {
                        var overPals = DyGetOverlappingDecadePalaces(palaces, lc.startAge, lc.endAge);
                        foreach (var (palName, _, _, _) in overPals)
                        {
                            string br = KbGetPalaceBranch(palaces, palName);
                            if (!string.IsNullOrEmpty(br)) uniqueDecadeBranches.Add(br);
                        }
                    }
                    foreach (var br in uniqueDecadeBranches)
                        decadeKbMap[br] = await KbZiweiQueryByBranch(ziweiPos, br);
                }

                string report = DyBuildReport(
                    yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch,
                    yStemSS, mStemSS, hStemSS, yBranchSS, mBranchSS, dBranchSS, hBranchSS,
                    dmElem, wuXing, bodyPct, bodyLabel, season, seaLabel,
                    pattern, yongShenElem, fuYiElem, yongReason, jiShenElem,
                    luckCycles, annualDetails, hasZiwei, palaces, siHuaDescMap,
                    ziweiFullContent, chartStars, decadeKbMap,
                    gender, birthYear, years, branches, dStem);

                // 九星氣學加成（純 KB）
                string dyNsSection = await NsBuildBirthSection(
                    user.BirthYear ?? birthYear,
                    user.BirthMonth ?? 1,
                    user.BirthDay ?? 1,
                    user.BirthHour ?? 0,
                    user.BirthGender ?? gender);
                if (!string.IsNullOrEmpty(dyNsSection)) report += dyNsSection;

                if (!dyIsAdmin) await RecordSubscriptionClaim(user.Id, dySubId, "BOOK_DAIYUN");
                string dyTitle = years == 0 ? "終身大運命書" : $"{years}年大運命書";
                await SaveUserReportAsync(user.Id, "daiyun", dyTitle, report,
                    new { years, birthYear = user.BirthYear, birthMonth = user.BirthMonth, birthDay = user.BirthDay, gender = user.BirthGender });
                return Ok(new { result = report, annualForecasts, baziTable, luckCycles = scoredCycles });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "大運命書失敗 User={User}", identity);
                return StatusCode(500, new { error = "大運命書生成失敗，請稍後再試", details = ex.Message });
            }
        }

        // === Lf Static Data Tables ===

        // 三干論斷（DB public."三干"，hardcoded）：key="三甲"/"四甲"...
        private static readonly Dictionary<string, string> LfSanGanMap = new()
        {
            ["三甲"]="天上貴，孤獨守空房", ["三乙"]="多陰私，又要敗祖業",
            ["三丙"]="人孤老，母在產中亡", ["三丁"]="多惡疾，手足也自傷",
            ["三戊"]="子隨出，離祖別家鄉", ["三己"]="別父母，兄弟各一方",
            ["三庚"]="是財郎，萬里置田莊", ["三辛"]="壽數長，財滯多災郎",
            ["三壬"]="家業盛，有富不久長", ["三癸"]="一亥全，烈火燒屋房",
            ["四甲"]="少夫妻", ["四乙"]="命早亡", ["四丙"]="子息空", ["四丁"]="壽不長",
            ["四戊"]="人孤刑", ["四己"]="人忠良", ["四庚"]="他鄉走", ["四辛"]="壽限長",
            ["四壬"]="定富足", ["四癸"]="人夭亡",
        };

        // 三支論斷（DB public."三支"，hardcoded）：key="三子"/"三丑"...
        private static readonly Dictionary<string, string> LfSanZhiMap = new()
        {
            ["三子"]="婚事重", ["三丑"]="四夫妻", ["三寅"]="守孤寡", ["三卯"]="兇惡多",
            ["三辰"]="好鬥傷", ["三巳"]="遭刑害", ["三午"]="克夫妻", ["三未"]="守空房",
            ["三申"]="人不足", ["三酉"]="獨居房", ["三戌"]="訟事多", ["三亥"]="孤苦憐",
        };

        // 地支神煞查表：地支(SKYNO) → 地支(TOFLO) → 神煞名稱
        // 來源：DB public."地支星剎"（四 KIND 資料相同）
        private static readonly Dictionary<string, Dictionary<string, string[]>> DiZhiShenShaMap = new()
        {
            ["子"] = new() { ["子"]=new[]{"將星"}, ["寅"]=new[]{"驛馬","孤辰"}, ["卯"]=new[]{"紅鸞"},
                             ["辰"]=new[]{"華蓋"}, ["巳"]=new[]{"劫煞"}, ["午"]=new[]{"災煞"},
                             ["酉"]=new[]{"桃花","天喜"}, ["戌"]=new[]{"寡宿"}, ["亥"]=new[]{"亡神"} },
            ["丑"] = new() { ["丑"]=new[]{"華蓋"}, ["寅"]=new[]{"劫煞","孤辰","紅鸞"}, ["卯"]=new[]{"災煞"},
                             ["午"]=new[]{"桃花"}, ["申"]=new[]{"亡神","天喜"}, ["酉"]=new[]{"將星"},
                             ["戌"]=new[]{"寡宿"}, ["亥"]=new[]{"驛馬"} },
            ["寅"] = new() { ["子"]=new[]{"災煞"}, ["丑"]=new[]{"寡宿","紅鸞"}, ["卯"]=new[]{"桃花"},
                             ["午"]=new[]{"將星"}, ["未"]=new[]{"天喜"}, ["申"]=new[]{"驛馬"},
                             ["巳"]=new[]{"亡神","孤辰"}, ["戌"]=new[]{"華蓋"}, ["亥"]=new[]{"劫煞"} },
            ["卯"] = new() { ["子"]=new[]{"桃花","紅鸞"}, ["丑"]=new[]{"寡宿"}, ["卯"]=new[]{"將星"},
                             ["午"]=new[]{"天喜"}, ["未"]=new[]{"華蓋"}, ["申"]=new[]{"劫煞"},
                             ["巳"]=new[]{"驛馬","孤辰"}, ["酉"]=new[]{"災煞"}, ["寅"]=new[]{"亡神"} },
            ["辰"] = new() { ["子"]=new[]{"將星"}, ["丑"]=new[]{"寡宿"}, ["辰"]=new[]{"華蓋"},
                             ["午"]=new[]{"災煞"}, ["酉"]=new[]{"桃花"}, ["寅"]=new[]{"驛馬"},
                             ["巳"]=new[]{"劫煞","孤辰","天喜"}, ["亥"]=new[]{"亡神","紅鸞"} },
            ["巳"] = new() { ["丑"]=new[]{"華蓋"}, ["寅"]=new[]{"劫煞"}, ["卯"]=new[]{"災煞"},
                             ["辰"]=new[]{"寡宿","天喜"}, ["午"]=new[]{"桃花"}, ["申"]=new[]{"亡神","孤辰"},
                             ["酉"]=new[]{"將星"}, ["戌"]=new[]{"紅鸞"}, ["亥"]=new[]{"驛馬"} },
            ["午"] = new() { ["子"]=new[]{"災煞"}, ["卯"]=new[]{"桃花","天喜"}, ["辰"]=new[]{"寡宿"},
                             ["午"]=new[]{"將星"}, ["巳"]=new[]{"亡神"}, ["申"]=new[]{"驛馬","孤辰"},
                             ["酉"]=new[]{"紅鸞"}, ["戌"]=new[]{"華蓋"}, ["亥"]=new[]{"劫煞"} },
            ["未"] = new() { ["子"]=new[]{"桃花"}, ["卯"]=new[]{"將星"}, ["辰"]=new[]{"寡宿"},
                             ["未"]=new[]{"華蓋"}, ["巳"]=new[]{"驛馬"}, ["申"]=new[]{"劫煞","孤辰","紅鸞"},
                             ["酉"]=new[]{"災煞"}, ["寅"]=new[]{"亡神","天喜"} },
            ["申"] = new() { ["子"]=new[]{"將星"}, ["丑"]=new[]{"天喜"}, ["辰"]=new[]{"華蓋"},
                             ["午"]=new[]{"災煞"}, ["未"]=new[]{"寡宿","紅鸞"}, ["巳"]=new[]{"劫煞"},
                             ["酉"]=new[]{"桃花"}, ["寅"]=new[]{"驛馬"}, ["亥"]=new[]{"亡神","孤辰"} },
            ["酉"] = new() { ["子"]=new[]{"天喜"}, ["丑"]=new[]{"華蓋"}, ["卯"]=new[]{"災煞"},
                             ["午"]=new[]{"桃花","紅鸞"}, ["未"]=new[]{"寡宿"}, ["申"]=new[]{"亡神"},
                             ["酉"]=new[]{"將星"}, ["寅"]=new[]{"劫煞"}, ["亥"]=new[]{"驛馬","孤辰"} },
            ["戌"] = new() { ["子"]=new[]{"災煞"}, ["卯"]=new[]{"桃花"}, ["午"]=new[]{"將星"},
                             ["未"]=new[]{"寡宿"}, ["巳"]=new[]{"亡神","紅鸞"}, ["申"]=new[]{"驛馬"},
                             ["戌"]=new[]{"華蓋"}, ["亥"]=new[]{"劫煞","孤辰","天喜"} },
            ["亥"] = new() { ["子"]=new[]{"桃花"}, ["卯"]=new[]{"將星"}, ["辰"]=new[]{"紅鸞"},
                             ["未"]=new[]{"華蓋"}, ["巳"]=new[]{"驛馬"}, ["申"]=new[]{"劫煞"},
                             ["酉"]=new[]{"災煞"}, ["戌"]=new[]{"寡宿","天喜"}, ["寅"]=new[]{"亡神","孤辰"} },
        };

        // 天干神煞查表：天干(SKYNO) → 地支(TOFLO) → 神煞名稱
        // 來源：DB public."天干星剎" KIND=日（各KIND資料幾乎相同）
        private static readonly Dictionary<string, Dictionary<string, string[]>> TianGanShenShaMap = new()
        {
            ["甲"] = new() { ["子"]=new[]{"學士"}, ["丑"]=new[]{"乙貴","血刃"}, ["寅"]=new[]{"干祿"},
                             ["卯"]=new[]{"羊刃","桃花"}, ["巳"]=new[]{"文昌"}, ["午"]=new[]{"紅艷"},
                             ["未"]=new[]{"乙貴"}, ["申"]=new[]{"路空"}, ["酉"]=new[]{"飛刃"}, ["亥"]=new[]{"桃花"} },
            ["乙"] = new() { ["子"]=new[]{"乙貴"}, ["寅"]=new[]{"血刃"}, ["卯"]=new[]{"干祿","桃花"},
                             ["辰"]=new[]{"羊刃"}, ["午"]=new[]{"紅艷"}, ["申"]=new[]{"乙貴"},
                             ["戌"]=new[]{"飛刃"}, ["亥"]=new[]{"學士","桃花"} },
            ["丙"] = new() { ["子"]=new[]{"飛刃","桃花"}, ["寅"]=new[]{"紅艷"}, ["卯"]=new[]{"學士"},
                             ["辰"]=new[]{"血刃"}, ["巳"]=new[]{"干祿"}, ["午"]=new[]{"羊刃"},
                             ["申"]=new[]{"文昌","桃花"}, ["酉"]=new[]{"乙貴"}, ["亥"]=new[]{"乙貴"} },
            ["丁"] = new() { ["子"]=new[]{"桃花"}, ["寅"]=new[]{"學士"}, ["巳"]=new[]{"血刃"},
                             ["午"]=new[]{"干祿"}, ["未"]=new[]{"紅艷","羊刃"}, ["申"]=new[]{"桃花"},
                             ["酉"]=new[]{"乙貴","文昌"}, ["亥"]=new[]{"乙貴"} },
            ["戊"] = new() { ["子"]=new[]{"飛刃"}, ["丑"]=new[]{"乙貴"}, ["卯"]=new[]{"桃花"},
                             ["辰"]=new[]{"紅艷","血刃"}, ["巳"]=new[]{"干祿"}, ["午"]=new[]{"羊刃","學士"},
                             ["未"]=new[]{"乙貴"}, ["申"]=new[]{"文昌"} },
            ["己"] = new() { ["子"]=new[]{"乙貴"}, ["丑"]=new[]{"飛刃"}, ["辰"]=new[]{"紅艷"},
                             ["巳"]=new[]{"學士","血刃"}, ["午"]=new[]{"干祿"}, ["未"]=new[]{"羊刃"},
                             ["申"]=new[]{"乙貴"}, ["酉"]=new[]{"文昌"}, ["戌"]=new[]{"桃花"} },
            ["庚"] = new() { ["丑"]=new[]{"乙貴"}, ["卯"]=new[]{"飛刃"}, ["巳"]=new[]{"桃花"},
                             ["午"]=new[]{"學士"}, ["未"]=new[]{"乙貴","血刃"}, ["申"]=new[]{"干祿"},
                             ["酉"]=new[]{"羊刃"}, ["戌"]=new[]{"紅艷"}, ["亥"]=new[]{"文昌","桃花"} },
            ["辛"] = new() { ["子"]=new[]{"文昌"}, ["寅"]=new[]{"乙貴"}, ["辰"]=new[]{"飛刃"},
                             ["巳"]=new[]{"學士"}, ["午"]=new[]{"乙貴"}, ["申"]=new[]{"血刃"},
                             ["酉"]=new[]{"干祿","紅艷"}, ["戌"]=new[]{"羊刃"} },
            ["壬"] = new() { ["子"]=new[]{"紅艷","羊刃"}, ["寅"]=new[]{"文昌"}, ["卯"]=new[]{"乙貴"},
                             ["巳"]=new[]{"乙貴"}, ["午"]=new[]{"飛刃","桃花"}, ["申"]=new[]{"學士"},
                             ["戌"]=new[]{"血刃"}, ["亥"]=new[]{"干祿"} },
            ["癸"] = new() { ["子"]=new[]{"干祿"}, ["丑"]=new[]{"羊刃"}, ["卯"]=new[]{"文昌","乙貴"},
                             ["巳"]=new[]{"乙貴"}, ["未"]=new[]{"飛刃"}, ["申"]=new[]{"紅艷","學士"},
                             ["亥"]=new[]{"血刃"} },
        };

        private static readonly Dictionary<string, List<(string stem, double ratio)>> LfBranchHiddenRatio = new()
        {
            { "子", new() { ("癸", 1.0) } },
            { "丑", new() { ("己", 0.5), ("癸", 0.3), ("辛", 0.2) } },
            { "寅", new() { ("甲", 0.5), ("丙", 0.3), ("戊", 0.2) } },
            { "卯", new() { ("乙", 1.0) } },
            { "辰", new() { ("戊", 0.5), ("乙", 0.3), ("癸", 0.2) } },
            { "巳", new() { ("丙", 0.5), ("戊", 0.3), ("庚", 0.2) } },
            { "午", new() { ("丁", 0.625), ("己", 0.375) } },
            { "未", new() { ("己", 0.5), ("丁", 0.3), ("乙", 0.2) } },
            { "申", new() { ("庚", 0.5), ("壬", 0.3), ("戊", 0.2) } },
            { "酉", new() { ("辛", 1.0) } },
            { "戌", new() { ("戊", 0.5), ("辛", 0.3), ("丁", 0.2) } },
            { "亥", new() { ("壬", 0.625), ("甲", 0.375) } }
        };

        // 生我 (印): element that generates day master
        private static readonly Dictionary<string, string> LfGenByElem = new()
            { { "木", "水" }, { "火", "木" }, { "土", "火" }, { "金", "土" }, { "水", "金" } };
        // 我生 (食傷): element day master generates
        private static readonly Dictionary<string, string> LfElemGen = new()
            { { "木", "火" }, { "火", "土" }, { "土", "金" }, { "金", "水" }, { "水", "木" } };
        // 我克 (財): element day master overcomes
        private static readonly Dictionary<string, string> LfElemOvercome = new()
            { { "木", "土" }, { "火", "金" }, { "土", "水" }, { "金", "木" }, { "水", "火" } };
        // 克我 (官殺): element that overcomes day master
        private static readonly Dictionary<string, string> LfElemOvercomeBy = new()
            { { "木", "金" }, { "火", "水" }, { "土", "木" }, { "金", "火" }, { "水", "土" } };

        // 地支主氣五行
        private static readonly Dictionary<string, string> LfBranchElem = new()
            { {"子","水"},{"丑","土"},{"寅","木"},{"卯","木"},{"辰","土"},{"巳","火"},
              {"午","火"},{"未","土"},{"申","金"},{"酉","金"},{"戌","土"},{"亥","水"} };
        // 地支生肖
        private static readonly Dictionary<string, string> LfBranchZodiac = new()
            { {"子","鼠"},{"丑","牛"},{"寅","虎"},{"卯","兔"},{"辰","龍"},{"巳","蛇"},
              {"午","馬"},{"未","羊"},{"申","猴"},{"酉","雞"},{"戌","狗"},{"亥","豬"} };
        // 地支六沖對應
        private static readonly Dictionary<string, string> LfBranchChongOf = new()
            { {"子","午"},{"午","子"},{"丑","未"},{"未","丑"},
              {"寅","申"},{"申","寅"},{"卯","酉"},{"酉","卯"},
              {"辰","戌"},{"戌","辰"},{"巳","亥"},{"亥","巳"} };

        private static readonly HashSet<string> LfChong = new()
            { "子午","午子","丑未","未丑","寅申","申寅","卯酉","酉卯","辰戌","戌辰","巳亥","亥巳" };

        private static readonly Dictionary<string, (string partner, string elem)> LfHe = new()
        {
            { "子", ("丑","土") }, { "丑", ("子","土") }, { "寅", ("亥","木") }, { "亥", ("寅","木") },
            { "卯", ("戌","火") }, { "戌", ("卯","火") }, { "辰", ("酉","金") }, { "酉", ("辰","金") },
            { "巳", ("申","水") }, { "申", ("巳","水") }, { "午", ("未","土") }, { "未", ("午","土") }
        };

        private static readonly List<(string[] branches, string elem)> LfSanHe = new()
        {
            (new[] { "申","子","辰" }, "水"), (new[] { "亥","卯","未" }, "木"),
            (new[] { "寅","午","戌" }, "火"), (new[] { "巳","酉","丑" }, "金")
        };

        private static readonly List<(string[] branches, string elem)> LfSanHui = new()
        {
            (new[] { "亥","子","丑" }, "水"), (new[] { "寅","卯","辰" }, "木"),
            (new[] { "巳","午","未" }, "火"), (new[] { "申","酉","戌" }, "金")
        };

        private static readonly HashSet<string> LfHai = new()
            { "子未","未子","丑午","午丑","寅巳","巳寅","卯辰","辰卯","申亥","亥申","酉戌","戌酉" };

        private static readonly HashSet<string> LfPo = new()
            { "子酉","酉子","丑辰","辰丑","寅亥","亥寅","卯午","午卯","申巳","巳申","戌未","未戌" };

        private static readonly List<string[]> LfXing = new()
            { new[] { "寅","巳","申" }, new[] { "丑","戌","未" }, new[] { "子","卯" } };

        // 二十四節氣人元司令用事表：月支 → [(藏干, 累積天數上限), ...]
        private static readonly Dictionary<string, (string stem, int endDay)[]> LfSilingTable = new()
        {
            ["寅"] = new[] { ("戊", 7),  ("丙", 14), ("甲", 30) },
            ["卯"] = new[] { ("甲", 10), ("乙", 30) },
            ["辰"] = new[] { ("乙", 9),  ("癸", 12), ("戊", 30) },
            ["巳"] = new[] { ("戊", 5),  ("庚", 14), ("丙", 30) },
            ["午"] = new[] { ("丙", 10), ("己", 19), ("丁", 30) },
            ["未"] = new[] { ("丁", 9),  ("乙", 12), ("己", 30) },
            ["申"] = new[] { ("戊", 10), ("壬", 13), ("庚", 30) },
            ["酉"] = new[] { ("庚", 10), ("辛", 30) },
            ["戌"] = new[] { ("辛", 9),  ("丁", 12), ("戊", 30) },
            ["亥"] = new[] { ("戊", 7),  ("甲", 12), ("壬", 30) },
            ["子"] = new[] { ("壬", 10), ("癸", 30) },
            ["丑"] = new[] { ("癸", 9),  ("辛", 12), ("己", 30) },
        };

        // 月支對應節氣名稱（每月兩個節氣）
        private static readonly Dictionary<string, (string jie, string qi)> LfBranchSolarTerms = new()
        {
            ["寅"] = ("立春", "雨水"),  ["卯"] = ("驚蟄", "春分"),  ["辰"] = ("清明", "穀雨"),
            ["巳"] = ("立夏", "小滿"),  ["午"] = ("芒種", "夏至"),  ["未"] = ("小暑", "大暑"),
            ["申"] = ("立秋", "處暑"),  ["酉"] = ("白露", "秋分"),  ["戌"] = ("寒露", "霜降"),
            ["亥"] = ("立冬", "小雪"),  ["子"] = ("大雪", "冬至"),  ["丑"] = ("小寒", "大寒"),
        };

        // 天干合 (五合)
        private static readonly Dictionary<string, (string stem, string elem)> LfTianGanHeMap = new()
        {
            { "甲", ("己", "土") }, { "己", ("甲", "土") },
            { "乙", ("庚", "金") }, { "庚", ("乙", "金") },
            { "丙", ("辛", "水") }, { "辛", ("丙", "水") },
            { "丁", ("壬", "木") }, { "壬", ("丁", "木") },
            { "戊", ("癸", "火") }, { "癸", ("戊", "火") },
        };

        private static double LfSeasonMult(string element, string season) => (element, season) switch
        {
            // 旺=1.6, 相=1.2, 休=1.0, 囚=0.5, 死=0.5
            ("木","春") => 1.6, ("木","夏") => 1.0, ("木","秋") => 0.5, ("木","冬") => 1.2, ("木","四季") => 0.5,
            ("火","春") => 1.2, ("火","夏") => 1.6, ("火","秋") => 0.5, ("火","冬") => 0.5, ("火","四季") => 1.0,
            ("土","春") => 0.5, ("土","夏") => 1.2, ("土","秋") => 1.0, ("土","冬") => 0.5, ("土","四季") => 1.6,
            ("金","春") => 0.5, ("金","夏") => 0.5, ("金","秋") => 1.6, ("金","冬") => 1.0, ("金","四季") => 1.2,
            ("水","春") => 1.0, ("水","夏") => 0.5, ("水","秋") => 1.2, ("水","冬") => 1.6, ("水","四季") => 0.5,
            _ => 1.0
        };

        // === Lf Pillar Helpers ===

        private static JsonElement LfGetPillar(JsonElement bazi, string key)
        {
            if (bazi.TryGetProperty(key, out var p)) return p;
            string cap = char.ToUpper(key[0]) + key[1..];
            if (bazi.TryGetProperty(cap, out p)) return p;
            return default;
        }

        private static string LfPillarStem(JsonElement p)
        {
            if (p.ValueKind == JsonValueKind.Undefined) return "";
            if (p.TryGetProperty("heavenlyStem", out var v) || p.TryGetProperty("HeavenlyStem", out v))
                return v.GetString() ?? "";
            return "";
        }

        private static string LfPillarBranch(JsonElement p)
        {
            if (p.ValueKind == JsonValueKind.Undefined) return "";
            if (p.TryGetProperty("earthlyBranch", out var v) || p.TryGetProperty("EarthlyBranch", out v))
                return v.GetString() ?? "";
            return "";
        }

        private static string LfPillarStemSS(JsonElement p)
        {
            if (p.ValueKind == JsonValueKind.Undefined) return "";
            if (p.TryGetProperty("heavenlyStemLiuShen", out var v) || p.TryGetProperty("HeavenlyStemLiuShen", out v))
                return v.GetString() ?? "";
            return "";
        }

        private static string LfPillarNaYin(JsonElement p)
        {
            if (p.ValueKind == JsonValueKind.Undefined) return "";
            if (p.TryGetProperty("naYin", out var v) || p.TryGetProperty("NaYin", out v))
                return v.GetString() ?? "";
            return "";
        }

        private static List<object> LfPillarHiddenPairs(JsonElement p)
        {
            var result = new List<object>();
            if (p.ValueKind == JsonValueKind.Undefined) return result;
            if (!p.TryGetProperty("hiddenStemLiuShen", out var arr) && !p.TryGetProperty("HiddenStemLiuShen", out arr))
                return result;
            if (arr.ValueKind != JsonValueKind.Array) return result;
            var items = arr.EnumerateArray().Select(x => x.GetString() ?? "").ToList();
            for (int i = 0; i + 1 < items.Count; i += 2)
                result.Add(new { ss = items[i], stem = items[i + 1] });
            return result;
        }

        private static string LfPillarBranchMainSS(JsonElement p)
        {
            if (p.ValueKind == JsonValueKind.Undefined) return "";
            if (p.TryGetProperty("hiddenStemLiuShen", out var arr) && arr.ValueKind == JsonValueKind.Array)
                return arr.EnumerateArray().FirstOrDefault().GetString() ?? "";
            return "";
        }

        private static List<(string stem, string branch, string liuShen, int startAge, int endAge)> LfExtractLuckCycles(JsonElement root)
        {
            var list = new List<(string, string, string, int, int)>();
            if (!root.TryGetProperty("baziLuckCycles", out var cycles) || cycles.ValueKind != JsonValueKind.Array)
                return list;
            foreach (var c in cycles.EnumerateArray())
            {
                string s  = c.TryGetProperty("heavenlyStem",  out var hs) ? hs.GetString() ?? "" : "";
                string b  = c.TryGetProperty("earthlyBranch", out var eb) ? eb.GetString() ?? "" : "";
                string ls = c.TryGetProperty("liuShen",       out var lv) ? lv.GetString() ?? "" : "";
                int sa    = c.TryGetProperty("startAge",      out var sa2) ? sa2.GetInt32() : 0;
                int ea    = c.TryGetProperty("endAge",        out var ea2) ? ea2.GetInt32() : 0;
                if (!string.IsNullOrEmpty(s)) list.Add((s, b, ls, sa, ea));
            }
            return list;
        }

        // === Lf Core Calculation Methods ===

        private static string LfGetSeason(string mb) => mb switch
        {
            "寅" or "卯" => "春", "巳" or "午" => "夏",
            "申" or "酉" => "秋", "亥" or "子" => "冬",
            _ => "四季"
        };

        private static string LfGetSeasonLabel(string mb) => mb switch
        {
            "亥" or "子" or "丑" => "寒凍", "巳" or "午" or "未" => "炎熱",
            "寅" or "卯" or "辰" => "濕木", _ => "燥金"
        };

        private static double LfGetBranchMult(string branch, string[] allBranches)
        {
            double m = 1.0;
            // 三會 highest priority
            foreach (var (brs, _) in LfSanHui)
                if (brs.Contains(branch) && brs.All(b => allBranches.Contains(b))) { m *= 2.5; break; }
            // 三合
            foreach (var (brs, _) in LfSanHe)
                if (brs.Contains(branch) && brs.All(b => allBranches.Contains(b))) { m *= 2.0; break; }
            // 六合
            if (LfHe.TryGetValue(branch, out var heI) && allBranches.Contains(heI.partner)) m *= 1.10;
            // 六沖
            if (allBranches.Where(b => b != branch).Any(b => LfChong.Contains(branch + b))) m *= 0.50;
            // 三刑
            foreach (var xg in LfXing)
                if (xg.Contains(branch) && xg.Count(x => allBranches.Contains(x)) >= 2) { m *= 0.60; break; }
            // 六害
            if (allBranches.Where(b => b != branch).Any(b => LfHai.Contains(branch + b))) m *= 0.70;
            // 六破
            if (allBranches.Where(b => b != branch).Any(b => LfPo.Contains(branch + b))) m *= 0.75;
            return m;
        }

        // 天干通根乘數：本氣(主氣)通根1.5，中氣/餘氣通根1.3，無根1.0
        private static double LfGetStemRootMult(string stemElem, string[] branches)
        {
            double best = 1.0;
            foreach (var branch in branches)
            {
                if (!LfBranchHiddenRatio.TryGetValue(branch, out var hidden)) continue;
                for (int i = 0; i < hidden.Count; i++)
                {
                    if (KbStemToElement(hidden[i].stem) == stemElem)
                    {
                        double m = i == 0 ? 1.5 : 1.3; // 本氣1.5，中氣/餘氣1.3
                        if (m > best) best = m;
                        break;
                    }
                }
            }
            return best;
        }

        private static Dictionary<string, double> LfCalcWuXingMatrix(
            string yStem, string yBranch, string mStem, string mBranch,
            string dStem, string dBranch, string hStem, string hBranch)
        {
            string season = LfGetSeason(mBranch);
            var scores = new Dictionary<string, double> { {"木",0},{"火",0},{"土",0},{"金",0},{"水",0} };
            var branches = new[] { yBranch, mBranch, dBranch, hBranch };

            void AddStem(string stem, double pts)
            {
                string elem = KbStemToElement(stem);
                if (!string.IsNullOrEmpty(elem))
                {
                    double rootMult = LfGetStemRootMult(elem, branches);
                    scores[elem] += pts * LfSeasonMult(elem, season) * rootMult;
                }
            }

            void AddBranch(string branch, double totalPts, double dispersionFactor = 1.0)
            {
                if (!LfBranchHiddenRatio.TryGetValue(branch, out var hidden)) return;
                double brMult = LfGetBranchMult(branch, branches);
                foreach (var (stem, ratio) in hidden)
                {
                    string elem = KbStemToElement(stem);
                    if (!string.IsNullOrEmpty(elem))
                        scores[elem] += totalPts * ratio * LfSeasonMult(elem, season) * brMult * dispersionFactor;
                }
            }

            // 月支藏干分散係數：藏干越多，主氣季節乘數效果適度降低
            // 1藏干(子/卯/酉等)=1.0，2藏干(亥/午)=0.85，3藏干(寅/申/巳/辰/戌/丑/未)=0.75
            // 待累積足夠案例後可再校正係數
            int mHiddenCount = LfBranchHiddenRatio.TryGetValue(mBranch, out var mHid) ? mHid.Count : 1;
            double mDispersion = mHiddenCount >= 3 ? 0.75 : mHiddenCount == 2 ? 0.85 : 1.0;

            AddStem(yStem, 10); AddStem(mStem, 10); AddStem(dStem, 10); AddStem(hStem, 10);
            AddBranch(yBranch, 10); AddBranch(dBranch, 10); AddBranch(hBranch, 10);
            AddBranch(mBranch, 30, mDispersion);

            double total = scores.Values.Sum();
            if (total > 0)
                foreach (var k in scores.Keys.ToList()) scores[k] = Math.Round(scores[k] / total * 100, 1);
            return scores;
        }

        private static double LfGetBodyStrengthPct(string dmElem, Dictionary<string, double> wuXing)
        {
            if (string.IsNullOrEmpty(dmElem)) return 50;
            string inElem   = LfGenByElem.GetValueOrDefault(dmElem, "");
            string outElem  = LfElemGen.GetValueOrDefault(dmElem, "");
            string caiElem  = LfElemOvercome.GetValueOrDefault(dmElem, "");
            string guanElem = LfElemOvercomeBy.GetValueOrDefault(dmElem, "");
            double biJi  = wuXing.GetValueOrDefault(dmElem, 0) + wuXing.GetValueOrDefault(inElem, 0);
            double xieKe = wuXing.GetValueOrDefault(outElem, 0) + wuXing.GetValueOrDefault(caiElem, 0) + wuXing.GetValueOrDefault(guanElem, 0);
            double total = biJi + xieKe;
            return total == 0 ? 50 : Math.Round(biJi / total * 100, 1);
        }

        private static string LfGetBodyStrengthLabel(double pct) => pct switch
        {
            >= 70 => "身強（極強）", >= 60 => "身強", >= 45 => "中和", >= 30 => "身弱", _ => "身弱（極弱）"
        };

        // 依格局 x 日主強弱 x 命局組合，取古文用神候選清單（優先序由前到後）
        private static string[] LfGetYongShenCandidates(
            string pattern, string dmElem, double bodyPct,
            Dictionary<string, double> wuXing)
        {
            string inElem   = LfGenByElem.GetValueOrDefault(dmElem, "");      // 印
            string biElem   = dmElem;                                           // 比劫
            string shiElem  = LfElemGen.GetValueOrDefault(dmElem, "");        // 食傷
            string caiElem  = LfElemOvercome.GetValueOrDefault(dmElem, "");   // 財
            string guanElem = LfElemOvercomeBy.GetValueOrDefault(dmElem, ""); // 官殺

            bool isStrong = bodyPct >= 60;
            bool isWeak   = bodyPct < 45;

            bool caiHeavy  = wuXing.GetValueOrDefault(caiElem, 0)  >= 15;
            bool shiHeavy  = wuXing.GetValueOrDefault(shiElem, 0)  >= 15;
            bool guanHeavy = wuXing.GetValueOrDefault(guanElem, 0) >= 15;
            bool inHeavy   = wuXing.GetValueOrDefault(inElem, 0)   >= 15;
            bool biHeavy   = wuXing.GetValueOrDefault(biElem, 0)   >= 15;

            return pattern switch
            {
                "正官格" when isWeak && caiHeavy  => new[] { biElem, inElem },
                "正官格" when isWeak && shiHeavy  => new[] { inElem, biElem },
                "正官格" when isWeak && guanHeavy => new[] { inElem },
                "正官格" when isWeak              => new[] { inElem, biElem },
                "正官格" when isStrong && biHeavy  => new[] { guanElem },
                "正官格" when isStrong && inHeavy  => new[] { caiElem },
                "正官格" when isStrong && shiHeavy => new[] { caiElem },
                "正官格"                           => new[] { guanElem, caiElem },

                "七殺格" when isWeak && caiHeavy  => new[] { biElem, inElem },
                "七殺格" when isWeak && shiHeavy  => new[] { inElem, biElem },
                "七殺格" when isWeak && guanHeavy => new[] { inElem },
                "七殺格" when isWeak              => new[] { inElem, biElem },
                "七殺格" when isStrong && biHeavy  => new[] { guanElem },
                "七殺格" when isStrong && inHeavy  => new[] { caiElem },
                "七殺格" when isStrong && guanHeavy => new[] { shiElem, caiElem },
                "七殺格"                            => new[] { guanElem, shiElem },

                "正財格" or "偏財格" when isWeak && shiHeavy  => new[] { inElem, biElem },
                "正財格" or "偏財格" when isWeak && caiHeavy  => new[] { biElem, inElem },
                "正財格" or "偏財格" when isWeak && guanHeavy => new[] { inElem, biElem },
                "正財格" or "偏財格" when isWeak              => new[] { biElem, inElem },
                "正財格" or "偏財格" when isStrong && biHeavy  => new[] { shiElem, guanElem },
                "正財格" or "偏財格" when isStrong && inHeavy  => new[] { caiElem },
                "正財格" or "偏財格"                           => new[] { guanElem, shiElem },

                "正印格" or "偏印格" when isWeak && guanHeavy => new[] { inElem },
                "正印格" or "偏印格" when isWeak && shiHeavy  => new[] { inElem },
                "正印格" or "偏印格" when isWeak && caiHeavy  => new[] { biElem, inElem },
                "正印格" or "偏印格" when isWeak              => new[] { inElem, biElem },
                "正印格" or "偏印格" when isStrong && biHeavy  => new[] { guanElem, shiElem },
                "正印格" or "偏印格" when isStrong && inHeavy  => new[] { caiElem },
                "正印格" or "偏印格" when isStrong && caiHeavy => new[] { guanElem },
                "正印格" or "偏印格"                           => new[] { guanElem, caiElem },

                "食神格" when isWeak && guanHeavy => new[] { inElem },
                "食神格" when isWeak && caiHeavy  => new[] { biElem, inElem },
                "食神格" when isWeak && shiHeavy  => new[] { inElem },
                "食神格" when isWeak              => new[] { inElem, biElem },
                "食神格" when isStrong && inHeavy  => new[] { caiElem },
                "食神格" when isStrong && biHeavy  => new[] { shiElem },
                "食神格" when isStrong && caiHeavy => new[] { guanElem },
                "食神格"                           => new[] { shiElem, caiElem },

                "傷官格" when isWeak && caiHeavy  => new[] { biElem, inElem },
                "傷官格" when isWeak && guanHeavy => new[] { inElem },
                "傷官格" when isWeak              => new[] { inElem, biElem },
                "傷官格" when isStrong && shiHeavy => new[] { inElem },
                "傷官格" when isStrong && biHeavy  => new[] { guanElem },
                "傷官格" when isStrong && inHeavy  => new[] { caiElem },
                "傷官格"                           => new[] { shiElem, caiElem },

                "建祿格" when isWeak && caiHeavy  => new[] { biElem },
                "建祿格" when isWeak && guanHeavy => new[] { inElem },
                "建祿格" when isWeak && shiHeavy  => new[] { inElem },
                "建祿格" when isWeak              => new[] { inElem, biElem },
                "建祿格" when isStrong && biHeavy  => new[] { guanElem },
                "建祿格" when isStrong && inHeavy  => new[] { caiElem },
                "建祿格" when isStrong && shiHeavy => new[] { caiElem },         // 古文：傷食多身強→財
                "建祿格" when isStrong && caiHeavy => new[] { guanElem, shiElem },
                "建祿格"                           => new[] { guanElem, shiElem },

                "月刃格" when caiHeavy  => new[] { guanElem },
                "月刃格" when guanHeavy => new[] { caiElem },
                "月刃格" when shiHeavy  => new[] { caiElem },
                "月刃格" when inHeavy   => new[] { caiElem },
                "月刃格" when biHeavy   => new[] { guanElem },
                "月刃格"                => new[] { guanElem, caiElem },

                _ when isStrong => new[] { shiElem, caiElem, guanElem },
                _               => new[] { inElem, biElem }
            };
        }

        // 內格三原則篩選用神（僅用於八格/建祿格/月刃格）
        // Step1: 有根優先；有2個以上有根 → Step2決勝
        // Step2: 月支旺相決勝（從有根集合 or 全候選中選旺→相）
        // Step3: 確認無沖剋合（有問題換下一候選重跑）
        private static string LfPickYongShen(
            string[] candidates,
            string yBranch, string mBranch, string dBranch, string hBranch,
            string[] chartStems, Dictionary<string, double> wuXing)
        {
            var allBranches = new[] { yBranch, mBranch, dBranch, hBranch };
            var (wang, xiang, _, _, _) = LfGetWangXiang(mBranch);

            // 內部：對候選集合執行 Step1+Step2，回傳最佳候選
            string PickFromPool(List<string> pool)
            {
                if (pool.Count == 0) return "";
                if (pool.Count == 1) return pool[0];
                // 多個 → 月支旺相決勝
                var wangHit  = pool.FirstOrDefault(e => e == wang);
                if (wangHit  != null) return wangHit;
                var xiangHit = pool.FirstOrDefault(e => e == xiang);
                if (xiangHit != null) return xiangHit;
                // 旺相都沒有 → 古文優先序第一位
                return pool[0];
            }

            // Step3：確認候選無沖剋合；若有問題換下一個
            string ValidateOrFallback(string picked)
            {
                if (string.IsNullOrEmpty(picked)) return candidates[0];
                // 若被合剋 → 嘗試下一個候選（只換一次）
                if (LfIsElemNeutralizedByChart(picked, chartStems, allBranches))
                {
                    var next = candidates.FirstOrDefault(e => e != picked
                        && !LfIsElemNeutralizedByChart(e, chartStems, allBranches));
                    if (next != null) return next;
                }
                return picked;
            }

            // Step1：有根集合
            var rooted = candidates
                .Where(e => LfElemHasRoot(e, yBranch, mBranch, dBranch, hBranch))
                .ToList();

            if (rooted.Count == 1)
                return ValidateOrFallback(rooted[0]);

            if (rooted.Count > 1)
                return ValidateOrFallback(PickFromPool(rooted));

            // Step2：皆無根 → 從全候選月支旺相決勝
            return ValidateOrFallback(PickFromPool(candidates.ToList()));
        }

        // 檢查某五行元素是否在地支有根（藏干包含該元素）
        private static bool LfElemHasRoot(string elem, string y, string m, string d, string h)
        {
            foreach (var b in new[] { y, m, d, h })
                if (LfBranchHiddenRatio.TryGetValue(b, out var hidden))
                    foreach (var (stem, _) in hidden)
                        if (KbStemToElement(stem) == elem) return true;
            return false;
        }

        private static string LfGetJiShenElem(string yongShenElem, string dmElem, double bodyPct, string pattern = "")
        {
            // 化氣格：忌神=克化神之元素（yongShenElem已為化神五行）
            if (LfHuaQiGeJuSet.Contains(pattern))
                return LfElemOvercomeBy.GetValueOrDefault(yongShenElem, "");
            // 五行從旺格/從旺格：忌神=克日干之元素（破格之神）
            if (pattern == "從旺格" || LfWuXingGeJuSet.Contains(pattern))
                return LfElemOvercomeBy.GetValueOrDefault(dmElem, "");
            // 從強格：忌神=印/比劫（不可幫身對抗旺勢）
            if (pattern == "從強格")
                return dmElem;  // 比劫（自身力量，逆旺勢）
            // 從殺/從財/從兒格：忌神=印（生日主使其有力量對抗旺勢，破格之神）
            if (pattern is "從殺格" or "從財格" or "從兒格")
                return LfGenByElem.GetValueOrDefault(dmElem, "");  // 印星（最大破格威脅）
            // 身弱（<45%）：大忌 = 克我（官殺），直接傷身
            // 身強（>=45%）：大忌 = 印星（生我讓身更旺，反被騙）
            // 中和（45-60%）：大忌依月令而定，暫以克身為主
            if (bodyPct < 45)
                return LfElemOvercomeBy.GetValueOrDefault(dmElem, "");
            return LfGenByElem.GetValueOrDefault(dmElem, "");
        }

        // exact Ten God using yin-yang distinction
        private static string LfStemShiShen(string stem, string dStem)
        {
            if (string.IsNullOrEmpty(stem) || string.IsNullOrEmpty(dStem)) return "";
            string sElem  = KbStemToElement(stem);
            string dmElem = KbStemToElement(dStem);
            if (string.IsNullOrEmpty(sElem) || string.IsNullOrEmpty(dmElem)) return "";
            bool sYang  = "甲丙戊庚壬".Contains(stem);
            bool dmYang = "甲丙戊庚壬".Contains(dStem);
            bool same   = sYang == dmYang;
            if (sElem == dmElem)  return same ? "比肩" : "劫財";
            if (sElem == LfGenByElem.GetValueOrDefault(dmElem, ""))    return same ? "偏印" : "正印";
            if (sElem == LfElemGen.GetValueOrDefault(dmElem, ""))      return same ? "食神" : "傷官";
            if (sElem == LfElemOvercome.GetValueOrDefault(dmElem, "")) return same ? "偏財" : "正財";
            if (sElem == LfElemOvercomeBy.GetValueOrDefault(dmElem, "")) return same ? "七殺" : "正官";
            return "";
        }

        // 五行從旺外格偵測（曲直/炎上/稼穡/從革/潤下）
        // Returns pattern name if matched, empty string otherwise.
        private static string LfDetectWuXingGeJu(
            string dmElem, string mBranch, string[] stems, string[] branches)
        {
            // Count how many branches from a given set appear in the chart
            int CountBranches(string[] set) => branches.Count(b => set.Contains(b));
            // Check if any forbidden stem or branch-element appears in the chart
            bool HasForbidden(string[] fStems, string[] fBranches) =>
                stems.Any(s => fStems.Contains(s)) || branches.Any(b => fBranches.Contains(b));

            return dmElem switch
            {
                // 曲直格: 甲乙日干, 生春月, 支全寅卯辰或亥卯未, 無庚辛申酉
                "木" when new[] { "寅","卯","辰" }.Contains(mBranch)
                    && (CountBranches(new[] { "寅","卯","辰" }) >= 2
                        || CountBranches(new[] { "亥","卯","未" }) >= 2)
                    && !HasForbidden(new[] { "庚","辛" }, new[] { "申","酉" })
                    => "曲直格",

                // 炎上格: 丙丁日干, 生夏月, 支全巳午未或寅午戌, 無壬癸亥子
                "火" when new[] { "巳","午","未" }.Contains(mBranch)
                    && (CountBranches(new[] { "巳","午","未" }) >= 2
                        || CountBranches(new[] { "寅","午","戌" }) >= 2)
                    && !HasForbidden(new[] { "壬","癸" }, new[] { "亥","子" })
                    => "炎上格",

                // 稼穡格: 戊己日干, 生四季月, 支全辰戌丑未, 無甲乙寅卯
                "土" when new[] { "辰","戌","丑","未" }.Contains(mBranch)
                    && CountBranches(new[] { "辰","戌","丑","未" }) >= 3
                    && !HasForbidden(new[] { "甲","乙" }, new[] { "寅","卯" })
                    => "稼穡格",

                // 從革格: 庚辛日干, 生秋月, 支全申酉戌或巳酉丑, 無丙丁午未
                "金" when new[] { "申","酉","戌" }.Contains(mBranch)
                    && (CountBranches(new[] { "申","酉","戌" }) >= 2
                        || CountBranches(new[] { "巳","酉","丑" }) >= 2)
                    && !HasForbidden(new[] { "丙","丁" }, new[] { "午","未" })
                    => "從革格",

                // 潤下格: 壬癸日干, 生冬月, 支全亥子丑或申子辰, 無戊己未戌
                "水" when new[] { "亥","子","丑" }.Contains(mBranch)
                    && (CountBranches(new[] { "亥","子","丑" }) >= 2
                        || CountBranches(new[] { "申","子","辰" }) >= 2)
                    && !HasForbidden(new[] { "戊","己" }, new[] { "未","戌" })
                    => "潤下格",

                _ => ""
            };
        }

        // 化氣格偵測：日干與月干或時干合化，月令支持化神，且無破格元素
        // allStems = 全部四柱天干（含日干），allBranches = 全部四柱地支
        private static (string pattern, string huaElem) LfDetectHuaQiGeJu(
            string dStem, string mStem, string hStem, string mBranch,
            string[] allStems, string[] allBranches)
        {
            foreach (var partner in new[] { mStem, hStem })
            {
                if (!LfHuaQiPairs.TryGetValue((dStem, partner), out var info)) continue;
                // 月令支持化神
                if (!info.mBranches.Contains(mBranch)) continue;
                // 無破格元素（排除日干本身與合化夥伴不計入禁忌天干）
                var otherStems = allStems.Where(s => s != dStem && s != partner).ToArray();
                if (otherStems.Any(s => info.forbidStems.Contains(s))) continue;
                if (allBranches.Any(b => info.forbidBranches.Contains(b))) continue;

                string patName = info.huaElem switch
                {
                    "土" => "化土格", "金" => "化金格", "水" => "化水格",
                    "木" => "化木格", "火" => "化火格", _ => ""
                };
                return (patName, info.huaElem);
            }
            return ("", "");
        }

        private static readonly HashSet<string> LfWuXingGeJuSet =
            new() { "曲直格", "炎上格", "稼穡格", "從革格", "潤下格" };

        private static readonly HashSet<string> LfHuaQiGeJuSet =
            new() { "化土格", "化金格", "化水格", "化木格", "化火格" };

        // 十干建祿月支對照（丙戊同祿在巳，丁己同祿在午）
        private static readonly Dictionary<string, string> LfJianLuBranch = new()
        { {"甲","寅"},{"乙","卯"},{"丙","巳"},{"戊","巳"},{"丁","午"},{"己","午"},
          {"庚","申"},{"辛","酉"},{"壬","亥"},{"癸","子"} };

        // 月刃月支對照（五陽干：甲卯/丙午/戊午/庚酉/壬子）
        private static readonly Dictionary<string, string> LfYueLanBranch = new()
        { {"甲","卯"},{"丙","午"},{"戊","午"},{"庚","酉"},{"壬","子"} };

        // 化氣格：(日干, 合化夥伴) → (化神五行, 月令支組, 破格天干, 破格地支)
        private static readonly Dictionary<(string, string), (string huaElem, string[] mBranches, string[] forbidStems, string[] forbidBranches)> LfHuaQiPairs = new()
        {
            { ("甲","己"), ("土", new[]{"辰","戌","丑","未"}, new[]{"甲","乙"},       new[]{"寅","卯"}) },
            { ("己","甲"), ("土", new[]{"辰","戌","丑","未"}, new[]{"甲","乙"},       new[]{"寅","卯"}) },
            { ("乙","庚"), ("金", new[]{"巳","酉","丑","申"}, new[]{"丙","丁"},       new[]{"巳","午"}) },
            { ("庚","乙"), ("金", new[]{"巳","酉","丑","申"}, new[]{"丙","丁"},       new[]{"巳","午"}) },
            { ("丙","辛"), ("水", new[]{"申","子","辰","亥"}, new[]{"戊","己"},       new[]{"辰","戌","丑","未"}) },
            { ("辛","丙"), ("水", new[]{"申","子","辰","亥"}, new[]{"戊","己"},       new[]{"辰","戌","丑","未"}) },
            { ("丁","壬"), ("木", new[]{"亥","卯","未","寅"}, new[]{"庚","辛"},       new[]{"申","酉"}) },
            { ("壬","丁"), ("木", new[]{"亥","卯","未","寅"}, new[]{"庚","辛"},       new[]{"申","酉"}) },
            { ("戊","癸"), ("火", new[]{"寅","午","戌","巳"}, new[]{"壬","癸"},       new[]{"亥","子"}) },
            { ("癸","戊"), ("火", new[]{"寅","午","戌","巳"}, new[]{"壬","癸"},       new[]{"亥","子"}) },
        };

        // 十天干調候用神表：dStem → mBranch → [調候干, 優先序從高到低]
        private static readonly Dictionary<string, Dictionary<string, string[]>> LfTiaoHou = new()
        {
            { "甲", new() {
                { "寅", new[] { "丙","癸" } }, { "卯", new[] { "庚","丙","丁" } }, { "辰", new[] { "庚","丁","壬" } },
                { "巳", new[] { "癸","丁","庚" } }, { "午", new[] { "癸","丁","庚" } }, { "未", new[] { "癸","丁","庚" } },
                { "申", new[] { "庚","丁","丙" } }, { "酉", new[] { "庚","丁","丙" } }, { "戌", new[] { "庚","丁","丙" } },
                { "亥", new[] { "丙" } }, { "子", new[] { "丙" } }, { "丑", new[] { "丙" } },
            }},
            { "乙", new() {
                { "寅", new[] { "丙","癸" } }, { "卯", new[] { "丙","癸" } }, { "辰", new[] { "癸","丙" } },
                { "巳", new[] { "癸","丙" } }, { "午", new[] { "癸","丙" } }, { "未", new[] { "癸","丙" } },
                { "申", new[] { "癸","丙" } }, { "酉", new[] { "癸","丙" } }, { "戌", new[] { "癸","丙" } },
                { "亥", new[] { "丙" } }, { "子", new[] { "丙" } }, { "丑", new[] { "丙" } },
            }},
            { "丙", new() {
                { "寅", new[] { "壬","庚" } }, { "卯", new[] { "壬","庚" } }, { "辰", new[] { "壬","甲" } },
                { "巳", new[] { "壬","庚" } }, { "午", new[] { "壬","庚" } }, { "未", new[] { "壬","庚" } },
                { "申", new[] { "壬","甲" } }, { "酉", new[] { "壬","甲" } }, { "戌", new[] { "壬","甲" } },
                { "亥", new[] { "甲","庚" } }, { "子", new[] { "甲","庚" } }, { "丑", new[] { "甲","庚" } },
            }},
            { "丁", new() {
                { "寅", new[] { "庚","甲" } }, { "卯", new[] { "庚","甲" } }, { "辰", new[] { "甲","庚" } },
                { "巳", new[] { "壬","庚" } }, { "午", new[] { "壬","庚" } }, { "未", new[] { "壬","庚" } },
                { "申", new[] { "甲","庚" } }, { "酉", new[] { "甲","庚" } }, { "戌", new[] { "甲","庚" } },
                { "亥", new[] { "甲","庚" } }, { "子", new[] { "甲","庚" } }, { "丑", new[] { "甲","庚" } },
            }},
            { "戊", new() {
                { "寅", new[] { "丙","甲","癸" } }, { "卯", new[] { "丙","甲","癸" } }, { "辰", new[] { "甲","丙","癸" } },
                { "巳", new[] { "癸","丙" } }, { "午", new[] { "癸","丙" } }, { "未", new[] { "癸","丙" } },
                { "申", new[] { "丙","癸" } }, { "酉", new[] { "丙","癸" } }, { "戌", new[] { "丙","癸" } },
                { "亥", new[] { "丙","甲" } }, { "子", new[] { "丙","甲" } }, { "丑", new[] { "丙","甲" } },
            }},
            { "己", new() {
                { "寅", new[] { "丙","庚" } }, { "卯", new[] { "甲","癸","丙" } }, { "辰", new[] { "丙","癸","甲" } },
                { "巳", new[] { "癸","丙" } }, { "午", new[] { "癸","丙" } }, { "未", new[] { "癸","丙" } },
                { "申", new[] { "丙","癸" } }, { "酉", new[] { "丙","癸" } }, { "戌", new[] { "丙","癸" } },
                { "亥", new[] { "丙","甲" } }, { "子", new[] { "丙","甲" } }, { "丑", new[] { "丙","甲" } },
            }},
            { "庚", new() {
                { "寅", new[] { "丙","甲","己" } }, { "卯", new[] { "丁","甲","庚" } }, { "辰", new[] { "甲","丁" } },
                { "巳", new[] { "壬","癸" } }, { "午", new[] { "壬","癸" } }, { "未", new[] { "壬","癸" } },
                { "申", new[] { "丁","甲" } }, { "酉", new[] { "丁","甲" } }, { "戌", new[] { "丁","甲" } },
                { "亥", new[] { "丙","丁" } }, { "子", new[] { "丙","丁" } }, { "丑", new[] { "丙","丁" } },
            }},
            { "辛", new() {
                { "寅", new[] { "己","壬","庚" } }, { "卯", new[] { "壬","甲" } }, { "辰", new[] { "壬","甲" } },
                { "巳", new[] { "壬","癸" } }, { "午", new[] { "壬","癸" } }, { "未", new[] { "壬","癸" } },
                { "申", new[] { "壬","甲" } }, { "酉", new[] { "壬","甲" } }, { "戌", new[] { "壬","甲" } },
                { "亥", new[] { "丙","壬" } }, { "子", new[] { "丙","壬" } }, { "丑", new[] { "丙","壬" } },
            }},
            { "壬", new() {
                { "寅", new[] { "庚","丙" } }, { "卯", new[] { "庚","辛" } }, { "辰", new[] { "甲","庚" } },
                { "巳", new[] { "辛","壬" } }, { "午", new[] { "辛","壬" } }, { "未", new[] { "辛","壬" } },
                { "申", new[] { "丁","甲" } }, { "酉", new[] { "丁","甲" } }, { "戌", new[] { "丁","甲" } },
                { "亥", new[] { "丙","丁" } }, { "子", new[] { "丙","丁" } }, { "丑", new[] { "丙","丁" } },
            }},
            { "癸", new() {
                { "寅", new[] { "辛","丙" } }, { "卯", new[] { "庚","辛" } }, { "辰", new[] { "丙","庚","甲" } },
                { "巳", new[] { "庚","辛" } }, { "午", new[] { "庚","辛" } }, { "未", new[] { "庚","辛" } },
                { "申", new[] { "丁","甲" } }, { "酉", new[] { "丁","甲" } }, { "戌", new[] { "丁","甲" } },
                { "亥", new[] { "丙","丁" } }, { "子", new[] { "丙","丁" } }, { "丑", new[] { "丙","丁" } },
            }},
        };

        // 計算天干在四柱地支中的根氣總分（本氣=3，中氣=2，餘氣=1）
        private static int LfStemRootScore(string stem, string[] allBranches)
        {
            int score = 0;
            string stemElem = KbStemToElement(stem);
            foreach (var branch in allBranches)
            {
                if (!LfBranchHiddenRatio.TryGetValue(branch, out var hidden)) continue;
                for (int i = 0; i < hidden.Count; i++)
                {
                    if (KbStemToElement(hidden[i].stem) == stemElem)
                    {
                        score += Math.Max(0, 3 - i); // 本氣=3, 中氣=2, 餘氣=1
                        break;
                    }
                }
            }
            return score;
        }

        private static (string pattern, string yongShenElem, string fuYiElem, string reason, string tiaoHouElem) LfDetectGeJuAndYongShen(
            string yStem, string yBranch, string mStem, string mBranch,
            string dStem, string dBranch, string hStem, string hBranch,
            string dmElem, Dictionary<string, double> wuXing, double bodyPct, string season)
        {
            // 取格優先順序：透干 → [調候優先 > 根氣最強] → 皆不透取根氣最強藏干 → 比劫改外格
            var allHeavenStems = new[] { yStem, mStem, hStem };
            var allBranches    = new[] { yBranch, mBranch, dBranch, hBranch };

            // 取得調候用神表（dStem × mBranch）
            string[] tiaoHouList = LfTiaoHou.TryGetValue(dStem, out var th1) && th1.TryGetValue(mBranch, out var th2)
                ? th2 : Array.Empty<string>();
            // 調候補用神元素（首位調候干的五行，供命書標注）
            string tiaoHouElem = tiaoHouList.Length > 0 ? KbStemToElement(tiaoHouList[0]) : "";

            // ── 取格優先順序：化氣格 → 外格 → 建祿格/月刃格 → 內格 ──

            // Step 0: 化氣格（最優先，一旦成立直接回傳，覆蓋所有其他格局）
            var allFourStems = new[] { yStem, mStem, dStem, hStem };
            var (huaPattern, huaElem) = LfDetectHuaQiGeJu(dStem, mStem, hStem, mBranch, allFourStems, allBranches);
            if (!string.IsNullOrEmpty(huaPattern))
            {
                string huaFuYi = LfGenByElem.GetValueOrDefault(huaElem, ""); // 生化神者為輔
                return (huaPattern, huaElem, huaFuYi, $"{huaPattern}（{dStem}化{huaElem}，順化神旺勢）", "");
            }

            // Step 1: 外格判定（優先，一旦成立直接使用，跳過內格）
            string pattern = "";

            // 五行從旺外格
            string wuXingGeJu = LfDetectWuXingGeJu(
                dmElem, mBranch, allHeavenStems, new[] { yBranch, mBranch, dBranch, hBranch });
            if (!string.IsNullOrEmpty(wuXingGeJu))
                pattern = wuXingGeJu;

            // 從格/從旺格
            if (string.IsNullOrEmpty(pattern) && bodyPct <= 20)
            {
                // 前置條件：天干有比劫/印通根 → 不能從格
                bool hasRootedBiJiYin = new[] { yStem, mStem, hStem }.Any(s =>
                {
                    string ss = LfStemShiShen(s, dStem);
                    if (ss != "比肩" && ss != "劫財" && ss != "正印" && ss != "偏印") return false;
                    string sElem = KbStemToElement(s);
                    return allBranches.Any(b =>
                        LfBranchHiddenRatio.TryGetValue(b, out var h) && h.Any(hh => KbStemToElement(hh.stem) == sElem));
                });

                if (!hasRootedBiJiYin)
                {
                    string guanElemD = LfElemOvercomeBy.GetValueOrDefault(dmElem, "");
                    string caiElemD  = LfElemOvercome.GetValueOrDefault(dmElem, "");
                    string shiElemD  = LfElemGen.GetValueOrDefault(dmElem, "");
                    double guanPctD  = wuXing.GetValueOrDefault(guanElemD, 0);
                    double caiPctD   = wuXing.GetValueOrDefault(caiElemD, 0);
                    double shiPctD   = wuXing.GetValueOrDefault(shiElemD, 0);
                    double oppPct    = guanPctD + caiPctD + shiPctD;
                    if (oppPct >= 70)
                    {
                        if      (guanPctD >= caiPctD && guanPctD >= shiPctD) pattern = "從殺格";
                        else if (caiPctD  >= guanPctD && caiPctD >= shiPctD) pattern = "從財格";
                        else                                                  pattern = "從兒格";
                    }
                }
            }
            if (string.IsNullOrEmpty(pattern) && bodyPct >= 80)
            {
                double sameElem = wuXing.GetValueOrDefault(dmElem, 0) + wuXing.GetValueOrDefault(LfGenByElem.GetValueOrDefault(dmElem, ""), 0);
                if (sameElem >= 75) pattern = "從旺格";
            }

            // Step 2 & 3: 外格未成立 → 建祿格/月刃格 or 內格
            string chosenStem = "";

            // 優先查十干祿表：建祿格（含戊借丙祿/己借丁祿）及月刃格
            if (string.IsNullOrEmpty(pattern))
            {
                if (LfYueLanBranch.TryGetValue(dStem, out var ylb) && ylb == mBranch)
                    pattern = "月刃格";
                else if (LfJianLuBranch.TryGetValue(dStem, out var jlb) && jlb == mBranch)
                    pattern = "建祿格";
            }

            if (string.IsNullOrEmpty(pattern) && LfBranchHiddenRatio.TryGetValue(mBranch, out var mH) && mH.Count > 0)
            {
                // 月令非比劫藏干
                var nonBiJieMH = mH
                    .Where(h => { var ss = LfStemShiShen(h.stem, dStem); return ss != "比肩" && ss != "劫財"; })
                    .ToList();

                if (nonBiJieMH.Count == 0)
                {
                    // 月令藏干全是比劫 → 建祿格/月刃格
                    chosenStem = mH[0].stem;
                }
                else
                {
                    // 內格取格：優先取透出（年/月/時干出現）的非比劫藏干，ratio 高者優先
                    // 若無透出則取主氣（ratio 最高非比劫）
                    var sortedMH = nonBiJieMH.OrderByDescending(h => h.ratio).ToList();
                    var transparentMH = sortedMH.Where(h => allHeavenStems.Contains(h.stem)).ToList();
                    chosenStem = transparentMH.Count > 0 ? transparentMH[0].stem : sortedMH[0].stem;
                }

                string chosenSS = LfStemShiShen(chosenStem, dStem);
                pattern = chosenSS switch
                {
                    "正官" => "正官格", "七殺" => "七殺格", "正印" => "正印格", "偏印" => "偏印格",
                    "正財" => "正財格", "偏財" => "偏財格", "食神" => "食神格", "傷官" => "傷官格",
                    "比肩" => "建祿格", "劫財" => "月刃格", _ => "普通格"
                };
            }

            string yongShenElem;
            string fuYiElem;
            string reason;

            // 外格（從旺/曲直/炎上等）：用神=日干本元素，直接設定
            if (pattern == "從旺格" || LfWuXingGeJuSet.Contains(pattern))
            {
                yongShenElem = dmElem;
                fuYiElem     = dmElem;
                reason       = pattern == "從旺格" ? "從旺格（順旺勢）" : $"{pattern}（五行純粹，順旺勢）";
            }
            // 從強格：取旺勢最強非日主元素
            else if (pattern == "從強格")
            {
                yongShenElem = new[] { LfElemGen.GetValueOrDefault(dmElem,""), LfElemOvercome.GetValueOrDefault(dmElem,""), LfElemOvercomeBy.GetValueOrDefault(dmElem,"") }
                    .OrderByDescending(e => wuXing.GetValueOrDefault(e, 0)).First();
                fuYiElem = yongShenElem;
                reason   = "從強格（順旺勢）";
            }
            // 八格 + 建祿格 + 月刃格：候選清單 → 三原則篩選
            else
            {
                string[] candidates = LfGetYongShenCandidates(pattern, dmElem, bodyPct, wuXing);
                yongShenElem = LfPickYongShen(candidates, yBranch, mBranch, dBranch, hBranch, allHeavenStems, wuXing);

                string lfYongRole = "";
                if (bodyPct >= 60) {
                    string _guanE = LfElemOvercomeBy.GetValueOrDefault(dmElem, "");
                    string _shiE  = LfElemGen.GetValueOrDefault(dmElem, "");
                    string _caiE  = LfElemOvercome.GetValueOrDefault(dmElem, "");
                    if (yongShenElem == _guanE)      lfYongRole = "官殺制旺身";
                    else if (yongShenElem == _shiE)  lfYongRole = "食傷洩耗日主";
                    else if (yongShenElem == _caiE)  lfYongRole = "財星耗洩";
                    else lfYongRole = "洩耗";
                } else if (bodyPct < 45) {
                    string _inE  = LfGenByElem.GetValueOrDefault(dmElem, "");
                    string _biE  = dmElem;
                    if (yongShenElem == _inE)       lfYongRole = "印星生扶";
                    else if (yongShenElem == _biE)  lfYongRole = "比劫助身";
                    else lfYongRole = "生扶";
                } else {
                    lfYongRole = "通關平衡";
                }
                reason = $"古文格局法（{(bodyPct >= 60 ? "身強" : bodyPct < 45 ? "身弱" : "")}{pattern}，{lfYongRole}）";

                // fuYiElem：候選清單第二位（排除主用神與忌神）
                string tempJiShen = LfGetJiShenElem(yongShenElem, dmElem, bodyPct, pattern);
                fuYiElem = candidates
                    .Where(e => e != yongShenElem && e != tempJiShen)
                    .FirstOrDefault() ?? yongShenElem;
            }

            if (string.IsNullOrEmpty(yongShenElem)) yongShenElem = dmElem;
            if (string.IsNullOrEmpty(fuYiElem)) fuYiElem = yongShenElem;
            // 從格覆寫：用神/輔神/原因依格局類型重設，不套用扶抑法
            if (pattern == "從殺格")
            {
                yongShenElem = LfElemOvercomeBy.GetValueOrDefault(dmElem, ""); // 七殺（旺勢主力，從之）
                fuYiElem     = LfElemOvercome.GetValueOrDefault(dmElem, "");   // 財（生殺，輔助）
                reason       = "從殺格（日主順從七殺旺勢，忌印星、比劫）";
            }
            else if (pattern == "從財格")
            {
                yongShenElem = LfElemOvercome.GetValueOrDefault(dmElem, "");   // 財（旺勢主力，從之）
                fuYiElem     = LfElemGen.GetValueOrDefault(dmElem, "");        // 食傷（生財，輔助）
                reason       = "從財格（日主順從財星旺勢，忌印星、比劫）";
            }
            else if (pattern == "從兒格")
            {
                yongShenElem = LfElemGen.GetValueOrDefault(dmElem, "");        // 食傷（旺勢主力，從之）
                fuYiElem     = LfElemOvercome.GetValueOrDefault(dmElem, "");   // 財（食傷生財，輔助）
                reason       = "從兒格（日主順從食傷旺勢，忌官殺、印星）";
            }
            // 調候共用（夏巳午未 / 冬亥子丑二季優先，春秋補注）
            if (!string.IsNullOrEmpty(tiaoHouElem))
            {
                string jiShenForTiao = LfGetJiShenElem(yongShenElem, dmElem, bodyPct, pattern);
                bool isSummerWinter  = new[] { "巳","午","未","亥","子","丑" }.Contains(mBranch);
                if (tiaoHouElem == yongShenElem)
                    reason += "（扶抑調候同功）";
                else if (tiaoHouElem != jiShenForTiao)
                {
                    if (isSummerWinter)
                    {
                        // 調候與用神不抵觸 → 共用：調候進 fuYiElem
                        if (fuYiElem == yongShenElem || fuYiElem == jiShenForTiao)
                            fuYiElem = tiaoHouElem;
                        reason += $"；調候{tiaoHouElem}共用";
                    }
                    else
                        reason += $"；調候補用：{tiaoHouElem}";
                }
                else
                    reason += $"；調候{tiaoHouElem}受限（與忌神同，不採用）";
            }
            return (pattern, yongShenElem, fuYiElem, reason, tiaoHouElem);
        }

        // 依格局+身強弱返回善運/惡運元素組（對應運限論則表）
        private static (string[] good, string[] bad) LfGetPatternLuckElems(
            string pattern, string yongShenElem, string fuYiElem, string dmElem, bool isBodyStrong)
        {
            string genDm = LfGenByElem.GetValueOrDefault(dmElem, "");    // 印 (生我)
            string dmGen = LfElemGen.GetValueOrDefault(dmElem, "");      // 食傷 (我生)
            string dmKe  = LfElemOvercome.GetValueOrDefault(dmElem, ""); // 財 (我克)
            string keDm  = LfElemOvercomeBy.GetValueOrDefault(dmElem, ""); // 官殺 (克我)

            return pattern switch
            {
                "曲直格" => (new[] { "水","木","火" },   new[] { "金" }),
                "炎上格" => (new[] { "木","火","土" },   new[] { "水" }),
                "稼穡格" => (new[] { "火","土","金" },   new[] { "木" }),
                "從革格" => (new[] { "土","金","水" },   new[] { "火" }),
                "潤下格" => (new[] { "金","水","木" },   new[] { "土" }),
                "從財格" => (new[] { dmGen, dmKe, keDm }, new[] { genDm, dmElem }),
                "從殺格" => (new[] { dmKe, keDm },        new[] { genDm, dmElem }),
                "從兒格" => (new[] { dmGen, dmKe },       new[] { keDm, genDm }),
                "從旺格" => (new[] { genDm, dmElem },     new[] { dmKe, keDm, dmGen }),
                // 從強格：日主順從旺勢（yongShenElem為旺元素），善=生旺+旺+旺所生，惡=克旺
                "從強格" => (
                    new[] {
                        LfGenByElem.GetValueOrDefault(yongShenElem, ""),
                        yongShenElem,
                        LfElemGen.GetValueOrDefault(yongShenElem, "")
                    }.Where(e => !string.IsNullOrEmpty(e)).Distinct().ToArray(),
                    new[] { LfElemOvercomeBy.GetValueOrDefault(yongShenElem, "") }
                        .Where(e => !string.IsNullOrEmpty(e)).ToArray()
                ),
                "化土格" => (new[] { "火","土","金" },   new[] { "木" }),
                "化金格" => (new[] { "土","金","水" },   new[] { "火" }),
                "化水格" => (new[] { "金","水","木" },   new[] { "土" }),
                "化木格" => (new[] { "水","木","火" },   new[] { "金" }),
                "化火格" => (new[] { "木","火","土" },   new[] { "水" }),
                // 普通格（正官/七殺/財/印/食神/傷官/建祿/月刃）依身強弱
                _ when isBodyStrong => (new[] { dmGen, dmKe, keDm }, new[] { genDm, dmElem }),
                _                   => (new[] { genDm, dmElem },     new[] { dmKe, keDm, dmGen })
            };
        }

        // 檢查元素是否被命局克去（克合修正：善惡效力減半）
        private static bool LfIsElemNeutralizedByChart(string elem, string[] chartStems, string[] chartBranches)
        {
            string overcomeElem = LfElemOvercomeBy.GetValueOrDefault(elem, "");
            if (string.IsNullOrEmpty(overcomeElem)) return false;
            foreach (var s in chartStems)
                if (KbStemToElement(s) == overcomeElem) return true;
            foreach (var b in chartBranches)
                if (LfBranchHiddenRatio.TryGetValue(b, out var bh) && bh.Count > 0 && KbStemToElement(bh[0].stem) == overcomeElem)
                    return true;
            return false;
        }

        // 大運/流年評分（方案A：格局善惡元素組 + 刑沖會合破疊加）
        private static int LfCalcLuckScore(
            string lcStem, string lcBranch,
            string pattern, string yongShenElem, string fuYiElem, string jiShenElem,
            string dmElem, bool isBodyStrong, string tiaoHouElem,
            string season, string[] chartBranches, string[] chartStems, string dStem)
        {
            var (goodElems, badElems) = LfGetPatternLuckElems(pattern, yongShenElem, fuYiElem, dmElem, isBodyStrong);
            // 調候補用：若未在善/惡組，加入善組
            if (!string.IsNullOrEmpty(tiaoHouElem) && !goodElems.Contains(tiaoHouElem) && !badElems.Contains(tiaoHouElem))
                goodElems = goodElems.Concat(new[] { tiaoHouElem }).ToArray();

            double score = 50.0;

            // 天干評分
            string lcStemElem = KbStemToElement(lcStem);
            if (!string.IsNullOrEmpty(lcStemElem))
            {
                double mult = LfIsElemNeutralizedByChart(lcStemElem, chartStems, chartBranches) ? 0.5 : 1.0;
                if (goodElems.Contains(lcStemElem))     score += 20 * mult;
                else if (badElems.Contains(lcStemElem)) score -= 20 * mult;
            }

            // 地支評分（依藏干比例）
            string lcBranchMainElem = "";
            if (LfBranchHiddenRatio.TryGetValue(lcBranch, out var lcBH))
            {
                if (lcBH.Count > 0) lcBranchMainElem = KbStemToElement(lcBH[0].stem);
                foreach (var (hstem, ratio) in lcBH)
                {
                    string e = KbStemToElement(hstem);
                    if (string.IsNullOrEmpty(e)) continue;
                    double mult = LfIsElemNeutralizedByChart(e, chartStems, chartBranches) ? 0.5 : 1.0;
                    if (goodElems.Contains(e))     score += 20 * ratio * mult;
                    else if (badElems.Contains(e)) score -= 20 * ratio * mult;
                }
            }

            // 調候補助：冬=火，夏=水（季節暖/涼調節）
            string tuneElem = season == "冬" ? "火" : season == "夏" ? "水" : "";
            if (!string.IsNullOrEmpty(tuneElem))
            {
                if (lcStemElem == tuneElem && !goodElems.Contains(tuneElem))
                    score += badElems.Contains(tuneElem) ? 6 : 8;
                if (!string.IsNullOrEmpty(lcBranchMainElem) && lcBranchMainElem == tuneElem
                    && lcBranchMainElem != lcStemElem && !goodElems.Contains(tuneElem))
                    score += badElems.Contains(tuneElem) ? 3 : 5;
            }

            // 刑沖會合破疊加評分
            foreach (var cb in chartBranches)
            {
                if (!LfChong.Contains(lcBranch + cb)) continue;
                string cbElem = LfBranchHiddenRatio.TryGetValue(cb, out var cbH) && cbH.Count > 0
                    ? KbStemToElement(cbH[0].stem) : "";
                if (badElems.Contains(lcBranchMainElem) && goodElems.Contains(cbElem))  score -= 6;
                else if (goodElems.Contains(lcBranchMainElem) && badElems.Contains(cbElem)) score += 4;
                else if (badElems.Contains(lcBranchMainElem) && badElems.Contains(cbElem))  score -= 3;
            }
            foreach (var (brs, elem) in LfSanHui)
                if (brs.Contains(lcBranch) && brs.Count(b => b != lcBranch && chartBranches.Contains(b)) == 2)
                { if (goodElems.Contains(elem)) score += 10; else if (badElems.Contains(elem)) score -= 10; }
            foreach (var (brs, elem) in LfSanHe)
                if (brs.Contains(lcBranch) && brs.Count(b => b != lcBranch && chartBranches.Contains(b)) == 2)
                { if (goodElems.Contains(elem)) score += 7; else if (badElems.Contains(elem)) score -= 7; }
            // 三合半合（命局已有1字，大運帶第2字引動）
            foreach (var (brs, elem) in LfSanHe)
                if (brs.Contains(lcBranch) && brs.Count(b => b != lcBranch && chartBranches.Contains(b)) == 1)
                {
                    if (badElems.Contains(lcBranchMainElem))      score += 4;  // 忌神被半合，減凶
                    else if (goodElems.Contains(lcBranchMainElem)) score -= 3;  // 喜神被半合，吉稍減
                }
            // 三會半合
            foreach (var (brs, elem) in LfSanHui)
                if (brs.Contains(lcBranch) && brs.Count(b => b != lcBranch && chartBranches.Contains(b)) == 1)
                {
                    if (badElems.Contains(lcBranchMainElem))      score += 5;  // 忌神被半會，減凶
                    else if (goodElems.Contains(lcBranchMainElem)) score -= 4;  // 喜神被半會，吉稍減
                }
            if (LfHe.TryGetValue(lcBranch, out var heInfo) && chartBranches.Contains(heInfo.partner))
            { if (goodElems.Contains(heInfo.elem)) score += 4; else if (badElems.Contains(heInfo.elem)) score -= 4; }
            foreach (var xg in LfXing)
                if (xg.Contains(lcBranch) && xg.Count(b => b != lcBranch && chartBranches.Contains(b)) >= 1)
                    score -= 4;
            if (chartBranches.Any(b => LfHai.Contains(lcBranch + b))) score -= 3;
            if (chartBranches.Any(b => LfPo.Contains(lcBranch + b))) score -= 2;
            // 天干合：大運天干被命局合住 → 善效力減弱 or 惡效力減弱
            if (LfTianGanHeMap.TryGetValue(lcStem, out var tgHe) && chartStems.Contains(tgHe.stem))
            {
                if (goodElems.Contains(lcStemElem))     score -= 5;
                else if (badElems.Contains(lcStemElem)) score += 5;
            }

            return (int)Math.Round(Math.Clamp(score, 0, 100));
        }

        private static string LfLuckLevel(int score) => score switch
        {
            >= 80 => "大吉運", >= 65 => "中吉運", >= 50 => "平吉運", >= 38 => "平運", >= 20 => "中凶運", _ => "大凶運"
        };

        // 大運干支 vs 四柱完整關係表（供走勢總覽驗算）
        // 十神白話（喜/忌神各一句）
        private static string LfSsWhiteTalk(string ss, bool isGood) => (ss.Trim(), isGood) switch
        {
            ("比肩", true)  => "同儕助力，合作順暢，人際廣結",
            ("比肩", false) => "比肩奪財，同儕競爭激烈，財星受阻",
            ("劫財", true)  => "朋友助力，共謀發展",
            ("劫財", false) => "比劫旺盛，財星受剋，耗財破財損",
            ("食神", true)  => "才思靈動，事業順暢，財源活絡",
            ("食神", false) => "食神洩身，精氣神耗散，難有大作為",
            ("傷官", true)  => "才氣發揮，創意突出，可突破格局",
            ("傷官", false) => "傷官旺，官位受損，口舌是非不斷",
            ("偏財", true)  => "意外財至，偏門有利，人緣廣結",
            ("偏財", false) => "財被比劫奪，財來財去，難以積累",
            ("正財", true)  => "財星有力，正當收益穩健，事業踏實",
            ("正財", false) => "財星受剋，財運不佳，宜謹慎理財",
            ("偏官", true)  => "七殺制身得用，有衝勁，偏財助勢",
            ("偏官", false) => "七殺旺剋身，壓力沉重，阻礙重重",
            ("正官", true)  => "官星有力，名位穩固，貴人扶持",
            ("正官", false) => "官星剋身，工作壓力大，上司阻礙",
            ("偏印", true)  => "偏印護身，智慧開展，技藝有成",
            ("偏印", false) => "梟印奪食，進取心受阻，固執保守",
            ("正印", true)  => "印星護身，貴人扶持，學習進步",
            ("正印", false) => "印旺壓制進取動能，格局受限",
            _              => isGood ? "喜神有力，吉象顯現" : "忌神作祟，凶象發動",
        };

        // 大運逐步吉凶驗算（評分分解 + 有影響關係 + 空亡 + 神煞 + 白話論斷）
        private static string LfDyStepVerifyStr(
            string lcStem, string lcBranch, string lcStemSS, string lcBranchSS,
            string[] chartBranches, string[] chartStems,
            string[] chartBranchSS, string[] chartStemSS,
            string[] goodElems, string[] badElems,
            string[] dayEmpty, string dStem,
            int finalScore, string finalLevel)
        {
            var sb2 = new StringBuilder();
            var bl = new[] { "年", "月", "日", "時" };
            string lcStemElem = KbStemToElement(lcStem);
            string lcBranchMainElem = LfBranchHiddenRatio.TryGetValue(lcBranch, out var lcBH) && lcBH.Count > 0
                ? KbStemToElement(lcBH[0].stem) : "";
            bool stemBad  = badElems.Contains(lcStemElem);
            bool stemGood = goodElems.Contains(lcStemElem);
            bool brBad    = badElems.Contains(lcBranchMainElem);
            bool brGood   = goodElems.Contains(lcBranchMainElem);

            // ── 評分分解 ──────────────────────────────────
            var scoreParts = new List<string> { "基準50" };
            // 天干
            double stemMult = LfIsElemNeutralizedByChart(lcStemElem, chartStems, chartBranches) ? 0.5 : 1.0;
            if (stemBad)       scoreParts.Add($"干{lcStem}({lcStemSS}·忌){(int)(-20*stemMult):+0;-0}");
            else if (stemGood) scoreParts.Add($"干{lcStem}({lcStemSS}·喜){(int)(+20*stemMult):+0;-0}");
            // 地支藏干
            if (LfBranchHiddenRatio.TryGetValue(lcBranch, out var lcBH2))
                foreach (var (hs, ratio) in lcBH2)
                {
                    string e = KbStemToElement(hs);
                    bool eBad = badElems.Contains(e), eGood = goodElems.Contains(e);
                    if (!eBad && !eGood) continue;
                    double m2 = LfIsElemNeutralizedByChart(e, chartStems, chartBranches) ? 0.5 : 1.0;
                    int adj = (int)Math.Round((eGood ? 1 : -1) * 20 * ratio * m2);
                    scoreParts.Add($"支{lcBranch}({LfStemShiShen(hs, dStem)}·{(eBad ? "忌" : "喜")}){adj:+0;-0}");
                }

            // 關係加減收集（用於評分分解 + 有影響關係列表）
            var relImpacts = new List<(string desc, int impact, string why)>();
            // 天干合
            if (LfTianGanHeMap.TryGetValue(lcStem, out var tgHe) && chartStems.Contains(tgHe.stem))
            {
                int pi = Array.IndexOf(chartStems, tgHe.stem);
                int imp = stemGood ? -5 : stemBad ? +5 : 0;
                if (imp != 0) relImpacts.Add(($"干({lcStem})合{bl[pi]}干{chartStems[pi]}({chartStemSS[pi]})",
                    imp, stemBad ? "忌干被合，凶減" : "喜干被合，吉稍減"));
            }
            // 六沖
            for (int ci = 0; ci < chartBranches.Length; ci++)
            {
                if (!LfChong.Contains(lcBranch + chartBranches[ci])) continue;
                string cbElem = LfBranchHiddenRatio.TryGetValue(chartBranches[ci], out var cbH3) && cbH3.Count > 0
                    ? KbStemToElement(cbH3[0].stem) : "";
                int imp = (brBad && goodElems.Contains(cbElem)) ? -6
                        : (brGood && badElems.Contains(cbElem)) ? +4
                        : (brBad && badElems.Contains(cbElem)) ? -3 : 0;
                string why = (brBad && goodElems.Contains(cbElem)) ? "忌沖喜，凶增"
                           : (brGood && badElems.Contains(cbElem)) ? "沖走忌，凶減"
                           : "雙忌沖，動盪";
                if (imp != 0) relImpacts.Add(($"支({lcBranch})沖{bl[ci]}支{chartBranches[ci]}({chartBranchSS[ci]})", imp, why));
            }
            // 三會全局
            foreach (var (brs, elem) in LfSanHui)
            {
                if (!brs.Contains(lcBranch)) continue;
                var pts = brs.Where(b => b != lcBranch && chartBranches.Contains(b)).ToList();
                if (pts.Count != 2) continue;
                int imp = goodElems.Contains(elem) ? +10 : badElems.Contains(elem) ? -10 : 0;
                if (imp != 0) relImpacts.Add(($"支({lcBranch})三會成{elem}局", imp, imp > 0 ? "喜神三會，大吉" : "忌神三會，大凶"));
            }
            // 三合全局
            foreach (var (brs, elem) in LfSanHe)
            {
                if (!brs.Contains(lcBranch)) continue;
                var pts = brs.Where(b => b != lcBranch && chartBranches.Contains(b)).ToList();
                if (pts.Count != 2) continue;
                int imp = goodElems.Contains(elem) ? +7 : badElems.Contains(elem) ? -7 : 0;
                if (imp != 0) relImpacts.Add(($"支({lcBranch})三合成{elem}局", imp, imp > 0 ? "喜神三合，增吉" : "忌神三合，增凶"));
            }
            // 三合半合
            foreach (var (brs, elem) in LfSanHe)
            {
                if (!brs.Contains(lcBranch)) continue;
                var pts = brs.Where(b => b != lcBranch && chartBranches.Contains(b)).ToList();
                if (pts.Count != 1) continue;
                int pi = Array.IndexOf(chartBranches, pts[0]);
                int imp = brBad ? +4 : brGood ? -3 : 0;
                if (imp != 0) relImpacts.Add(($"支({lcBranch})半合{bl[pi]}支{pts[0]}({chartBranchSS[pi]})→{elem}局缺一字",
                    imp, brBad ? "忌神半合，凶減" : "喜神半合，吉稍減"));
            }
            // 三會半合
            foreach (var (brs, elem) in LfSanHui)
            {
                if (!brs.Contains(lcBranch)) continue;
                var pts = brs.Where(b => b != lcBranch && chartBranches.Contains(b)).ToList();
                if (pts.Count != 1) continue;
                int pi = Array.IndexOf(chartBranches, pts[0]);
                int imp = brBad ? +5 : brGood ? -4 : 0;
                if (imp != 0) relImpacts.Add(($"支({lcBranch})半會{bl[pi]}支{pts[0]}({chartBranchSS[pi]})→{elem}局缺一字",
                    imp, brBad ? "忌神半會，凶減" : "喜神半會，吉稍減"));
            }
            // 六合
            if (LfHe.TryGetValue(lcBranch, out var heInfo) && chartBranches.Contains(heInfo.partner))
            {
                int pi = Array.IndexOf(chartBranches, heInfo.partner);
                int imp = goodElems.Contains(heInfo.elem) ? +4 : badElems.Contains(heInfo.elem) ? -4 : 0;
                if (imp != 0) relImpacts.Add(($"支({lcBranch})六合{bl[pi]}支{heInfo.partner}({chartBranchSS[pi]})→{heInfo.elem}",
                    imp, brBad ? "忌神六合，凶減" : "喜神六合，吉稍減"));
            }
            // 三刑
            foreach (var xg in LfXing)
            {
                if (!xg.Contains(lcBranch)) continue;
                var pts = xg.Where(b => b != lcBranch && chartBranches.Contains(b)).ToList();
                if (pts.Count > 0) relImpacts.Add(($"支({lcBranch})刑{string.Join(",", pts)}", -4, "三刑，動盪損耗"));
            }
            // 六害
            for (int ci = 0; ci < chartBranches.Length; ci++)
                if (LfHai.Contains(lcBranch + chartBranches[ci]))
                    relImpacts.Add(($"支({lcBranch})害{bl[ci]}支{chartBranches[ci]}({chartBranchSS[ci]})", -3, "六害，暗中耗損"));

            // 輸出評分分解（含關係加減）
            foreach (var (desc, imp, _) in relImpacts)
                scoreParts.Add($"{desc}[{(imp >= 0 ? "+" : "")}{imp}]");
            sb2.AppendLine($"  評分分解：{string.Join(" ", scoreParts)} ＝ {finalScore}（{finalLevel}）");

            // ── 有影響的關係（跳過「無」）─────────────────
            if (relImpacts.Count > 0)
            {
                var relDescs = relImpacts.Select(r => $"{r.desc}→{r.why}[{(r.impact >= 0 ? "+" : "")}{r.impact}]");
                sb2.AppendLine($"  有影響關係：{string.Join("；", relDescs)}");
            }

            // ── 空亡 ───────────────────────────────────────
            string emptyStr = dayEmpty.Length >= 2 ? $"{dayEmpty[0]}{dayEmpty[1]}" : "無";
            bool brInEmpty  = dayEmpty.Length >= 2 && dayEmpty.Contains(lcBranch);
            bool stemInEmpty = dayEmpty.Length >= 2 && dayEmpty.Contains(lcBranch); // 同地支判斷
            string emptyDesc = brInEmpty
                ? (brBad ? $"運支{lcBranch}落旬空 → 忌神落空，凶力大減（空亡解凶）"
                          : $"運支{lcBranch}落旬空 → 喜神落空，吉力減半（空亡損吉）")
                : $"運支{lcBranch}不在旬空，{(brBad ? "忌神凶力正常發揮" : brGood ? "喜神吉力正常發揮" : "中性")}";
            sb2.AppendLine($"  空亡（日柱旬空：{emptyStr}）：{emptyDesc}");

            // ── 神煞（大運干支為基準，引動四柱地支）──────
            var shenShaHits = new List<string>();
            if (DiZhiShenShaMap.TryGetValue(lcBranch, out var dzRun))
                for (int pi = 0; pi < chartBranches.Length; pi++)
                    if (dzRun.TryGetValue(chartBranches[pi], out var ssArr))
                        foreach (var s in ssArr)
                            shenShaHits.Add($"{s}（運支→{bl[pi]}支，{LfShenShaStarDesc(s)}）");
            if (TianGanShenShaMap.TryGetValue(lcStem, out var tgRun))
                for (int pi = 0; pi < chartBranches.Length; pi++)
                    if (tgRun.TryGetValue(chartBranches[pi], out var ssArr))
                        foreach (var s in ssArr)
                            shenShaHits.Add($"{s}（運干→{bl[pi]}支，{LfShenShaStarDesc(s)}）");
            sb2.AppendLine(shenShaHits.Count > 0
                ? $"  神煞：{string.Join("、", shenShaHits)}"
                : "  神煞：無");

            // ── 白話論斷 ───────────────────────────────────
            string stemTalk = stemBad
                ? $"天干{lcStem}（{lcStemSS}）忌神掌事，{LfSsWhiteTalk(lcStemSS, false)}"
                : stemGood
                ? $"天干{lcStem}（{lcStemSS}）喜神得令，{LfSsWhiteTalk(lcStemSS, true)}"
                : $"天干{lcStem}（{lcStemSS}）中性，影響有限";
            string brTalk = brBad
                ? $"地支{lcBranch}（{lcBranchSS}）忌神掌事，{LfSsWhiteTalk(lcBranchSS, false)}"
                : brGood
                ? $"地支{lcBranch}（{lcBranchSS}）喜神得令，{LfSsWhiteTalk(lcBranchSS, true)}"
                : $"地支{lcBranch}（{lcBranchSS}）中性，影響有限";
            string relTalk = relImpacts.Count > 0
                ? "受干支關係牽制，" + string.Join("，",
                    relImpacts.Where(r => r.impact > 0).Select(r => r.why).Distinct().Take(2)) +
                  (relImpacts.Any(r => r.impact > 0) ? "，凶象有所緩和。" : "")
                : "";
            string emptyTalk = brInEmpty
                ? (brBad ? "忌神落旬空，凶力大減，是難得緩解之機。" : "喜神落空，吉力受損，宜謹慎。")
                : "";
            string overallTalk = finalScore >= 65 ? "整體走吉，宜積極進取，把握機遇。"
                : finalScore >= 50 ? "整體平穩偏吉，宜順勢而為，穩中求進。"
                : finalScore >= 38 ? "整體平偏凶，宜保守行事，避免大決策。"
                : "整體走凶，宜低調守成，謹防財務人事損耗。";
            sb2.AppendLine($"  論斷：{stemTalk}；{brTalk}。{relTalk}{emptyTalk}{overallTalk}");

            return sb2.ToString().TrimEnd();
        }


        // 偵測大運地支與命局地支的沖刑會合破，僅供文字描述，不影響評分
        private static string LfBranchEvents(string lcBranch, string[] chartBranches)
        {
            var events = new List<string>();
            // 六沖
            foreach (var cb in chartBranches)
                if (LfChong.Contains(lcBranch + cb)) events.Add($"與命局{cb}六沖");
            // 三會
            foreach (var (brs, elem) in LfSanHui)
                if (brs.Contains(lcBranch) && brs.Count(b => b != lcBranch && chartBranches.Contains(b)) == 2)
                    events.Add($"三會{elem}局");
            // 三合
            foreach (var (brs, elem) in LfSanHe)
                if (brs.Contains(lcBranch) && brs.Count(b => b != lcBranch && chartBranches.Contains(b)) == 2)
                    events.Add($"三合{elem}局");
            // 六合
            if (LfHe.TryGetValue(lcBranch, out var heInfo) && chartBranches.Contains(heInfo.partner))
                events.Add($"與{heInfo.partner}六合{heInfo.elem}");
            // 三刑
            foreach (var xg in LfXing)
                if (xg.Contains(lcBranch) && xg.Count(b => b != lcBranch && chartBranches.Contains(b)) >= 1)
                    events.Add("三刑");
            // 六害
            if (chartBranches.Any(b => LfHai.Contains(lcBranch + b))) events.Add("六害");
            // 六破
            if (chartBranches.Any(b => LfPo.Contains(lcBranch + b))) events.Add("六破");
            return events.Count > 0 ? string.Join("、", events) : "";
        }

        // ─── 年齡適切性過濾（共用，所有命書/大運/流年均適用）────────────────────────

        /// <summary>
        /// 判斷宮位論斷是否應依年齡跳過（年齡不符現實情境）
        /// </summary>
        private static bool LfShouldSkipPalace(string palace, int age)
        {
            if (age <= 0) return false;
            return palace switch
            {
                "父母宮" => age >= 65,              // 65+ 父母多已故世
                "夫妻宮" => age < 16 || age >= 75,  // 太小或75+ 不談婚戀
                "子女宮" => age < 18,               // 18前不談子女
                "官祿宮" => age >= 80,              // 80+ 不談事業升遷
                _ => false
            };
        }

        /// <summary>
        /// 依年齡調整宮位顯示名稱（如父母宮在61-64歲顯示為長輩宮）
        /// </summary>
        private static string LfPalaceAgeLabel(string palace, int age)
        {
            if (palace == "父母宮" && age >= 61 && age < 65) return "長輩宮";
            if (palace == "子女宮" && age >= 50) return "子女孫輩宮";
            return palace;
        }

        /// <summary>
        /// 依年齡返回論斷主題提示（適用於所有命書的章節開頭 / Gemini prompt）
        /// </summary>
        public static string LfAgeTopicHint(int age)
        {
            if (age <= 0) return "";
            return age switch
            {
                <= 15 => "【年齡提示】命主年幼，論斷著重學業、父母庇蔭、品格養成，不涉婚姻子女。",
                <= 25 => "【年齡提示】命主年輕，著重事業起步、感情萌芽。",
                <= 60 => "",
                <= 70 => "【年齡提示】命主已達退休年齡，父母多已不在，著重健康、財庫穩定、子女孫輩關係。",
                <= 80 => "【年齡提示】命主年事已高，不談父母（多已故），不談新感情，著重健康長壽、財庫守護、子女孫侍奉。",
                _    => "【年齡提示】命主高齡，著重健康養生、財庫守護、子女孫輩，其他宮位論斷從略。"
            };
        }

        // 大運地支與命局的宮位影響描述（含六神標記，替代原 LfBranchEvents 用於 Ch.10）
        private static string LfBranchEventsPalace(
            string lcBranch, string lcBranchSS,
            string[] chartBranches, string[] branchSS, int age = 0)
        {
            var lines = new List<string>();
            string[] palaceNames = { "父母宮", "兄弟宮", "夫妻宮", "子女宮" };

            // 逐支檢查：六沖/六合/三刑/六害/六破
            for (int i = 0; i < chartBranches.Length && i < palaceNames.Length; i++)
            {
                string cb = chartBranches[i];
                if (string.IsNullOrEmpty(cb) || cb == lcBranch) continue;

                string palace = palaceNames[i];
                if (LfShouldSkipPalace(palace, age)) continue; // 年齡過濾

                string cbSS = i < branchSS.Length ? (branchSS[i] ?? "") : "";
                string palaceLabel = LfPalaceAgeLabel(palace, age);
                string cbLabel = $"{palaceLabel}（{cbSS}·{cb}）";

                var rels = new List<string>();
                if (LfChong.Contains(lcBranch + cb))
                    rels.Add("六沖");
                if (LfHe.TryGetValue(lcBranch, out var heInfo) && heInfo.partner == cb)
                    rels.Add($"六合{heInfo.elem}局");
                foreach (var xg in LfXing)
                    if (xg.Contains(lcBranch) && xg.Contains(cb))
                    { rels.Add("相刑"); break; }
                if (LfHai.Contains(lcBranch + cb))
                    rels.Add("六害");
                if (LfPo.Contains(lcBranch + cb))
                    rels.Add("六破");

                if (rels.Count > 0)
                {
                    string impact = LfPalaceImpact(rels[0], palace, age);
                    lines.Add($"與{cbLabel}{string.Join("、", rels)}，{impact}");
                }
            }

            // 三會（需命局兩支配合）
            foreach (var (brs, elem) in LfSanHui)
            {
                if (!brs.Contains(lcBranch)) continue;
                var natals = brs.Where(b => b != lcBranch && chartBranches.Contains(b)).ToList();
                if (natals.Count == 2)
                {
                    var parts = natals
                        .Select(b => {
                            int idx = Array.IndexOf(chartBranches, b);
                            string ss  = idx >= 0 && idx < branchSS.Length    ? branchSS[idx]    : "";
                            string pal = idx >= 0 && idx < palaceNames.Length ? palaceNames[idx] : "";
                            return (pal, ss, b);
                        })
                        .Where(x => !LfShouldSkipPalace(x.pal, age))
                        .Select(x => $"{LfPalaceAgeLabel(x.pal, age)}（{x.ss}·{x.b}）")
                        .ToList();
                    string palDesc = parts.Count > 0 ? $"動及{string.Join("與", parts)}，" : "";
                    lines.Add($"三會{elem}局，{palDesc}此運{elem}氣大旺。");
                }
            }

            // 三合（需命局兩支配合）
            foreach (var (brs, elem) in LfSanHe)
            {
                if (!brs.Contains(lcBranch)) continue;
                var natals = brs.Where(b => b != lcBranch && chartBranches.Contains(b)).ToList();
                if (natals.Count == 2)
                {
                    var parts = natals
                        .Select(b => {
                            int idx = Array.IndexOf(chartBranches, b);
                            string ss  = idx >= 0 && idx < branchSS.Length    ? branchSS[idx]    : "";
                            string pal = idx >= 0 && idx < palaceNames.Length ? palaceNames[idx] : "";
                            return (pal, ss, b);
                        })
                        .Where(x => !LfShouldSkipPalace(x.pal, age))
                        .Select(x => $"{LfPalaceAgeLabel(x.pal, age)}（{x.ss}·{x.b}）")
                        .ToList();
                    string palDesc = parts.Count > 0 ? $"動及{string.Join("與", parts)}，" : "";
                    lines.Add($"三合{elem}局，{palDesc}此運{elem}氣得助。");
                }
            }

            return lines.Count > 0 ? string.Join("\n  ", lines) : "";
        }

        private static string LfPalaceImpact(string relType, string palace, int age = 0)
        {
            bool isBad = relType.Contains("沖") || relType.Contains("刑") || relType.Contains("害") || relType.Contains("破");
            string elderLabel = age >= 61 ? "長輩前輩" : "父母長輩";
            string childLabel = age >= 50 ? "子女孫輩" : "子女晚輩";
            return palace switch
            {
                "父母宮" => isBad ? $"留意{elderLabel}健康或緣分波動。" : $"{elderLabel}緣分加深，有貴人提攜。",
                "兄弟宮" => isBad ? "兄弟、友人、同事易有摩擦或分離。" : "兄弟、友人有助力，合作可期。",
                "夫妻宮" => isBad ? "配偶、感情易有波動，婚姻宜多溝通。" : "配偶緣分加深，感情婚姻有進展。",
                "子女宮" => isBad ? $"{childLabel}易有狀況，宜多關心。" : $"{childLabel}緣分佳，有喜訊。",
                _ => isBad ? "宜留意相關人事變動。" : "宜把握相關人事機緣。"
            };
        }

        private static int LfCalcRelativeScore(
            string stemSS, string branchSS, string yongShenElem, string jiShenElem,
            string pillarStem, string pillarBranch, string[] allBranches)
        {
            int score = 70;
            string stemElem = KbStemToElement(pillarStem);
            if (stemElem == yongShenElem) score += 15;
            else if (stemElem == jiShenElem) score -= 15;
            string brMainElem = LfBranchHiddenRatio.TryGetValue(pillarBranch, out var bh) && bh.Count > 0
                ? KbStemToElement(bh[0].stem) : "";
            if (brMainElem == yongShenElem) score += 10;
            else if (brMainElem == jiShenElem) score -= 10;
            if (LfHe.TryGetValue(pillarBranch, out var heI) && allBranches.Contains(heI.partner)) score += 10;
            foreach (var (brs, elem) in LfSanHui)
                if (brs.Contains(pillarBranch) && brs.Count(b => b != pillarBranch && allBranches.Contains(b)) >= 2 && elem == yongShenElem) score += 20;
            foreach (var (brs, elem) in LfSanHe)
                if (brs.Contains(pillarBranch) && brs.Count(b => b != pillarBranch && allBranches.Contains(b)) >= 2 && elem == yongShenElem) score += 15;
            if (allBranches.Where(b => b != pillarBranch).Any(b => LfChong.Contains(pillarBranch + b))) score -= 18;
            foreach (var xg in LfXing)
                if (xg.Contains(pillarBranch) && xg.Count(b => allBranches.Contains(b)) >= 2) { score -= 15; break; }
            if (allBranches.Where(b => b != pillarBranch).Any(b => LfHai.Contains(pillarBranch + b))) score -= 10;
            if (allBranches.Where(b => b != pillarBranch).Any(b => LfPo.Contains(pillarBranch + b))) score -= 8;
            return Math.Clamp(score, 0, 100);
        }

        private static string LfRelativeLevel(int score) => score switch
        {
            >= 80 => "緣深", >= 60 => "緣中", >= 40 => "緣薄", >= 20 => "緣弱", _ => "無緣"
        };

        private static int LfChineseNumToInt(string s) => s switch
        {
            "一" => 1, "二" => 2, "三" => 3, "四" => 4, "五" => 5,
            "六" => 6, "七" => 7, "八" => 8, "九" => 9, "十" => 10,
            "十一" => 11, "十二" => 12, _ => 99
        };

        // ─── Ch.5 六親論斷輔助方法 ─────────────────────────────────────────────

        private static bool LfRelHasRoot(string pillarStem, string branch)
        {
            string stemElem = KbStemToElement(pillarStem);
            return LfBranchHiddenRatio.TryGetValue(branch, out var hidden)
                && hidden.Any(h => KbStemToElement(h.stem) == stemElem);
        }

        // 地支本氣對天干的作用：生/剋/洩/財/比
        private static string LfRelBranchEffect(string branch, string pillarStem)
        {
            string stemElem = KbStemToElement(pillarStem);
            if (!LfBranchHiddenRatio.TryGetValue(branch, out var hidden) || hidden.Count == 0) return "比";
            string brElem = KbStemToElement(hidden[0].stem);
            if (brElem == stemElem)                                        return "比";
            if (LfElemGen.GetValueOrDefault(brElem) == stemElem)           return "生";
            if (LfElemOvercome.GetValueOrDefault(brElem) == stemElem)      return "剋";
            if (LfElemGen.GetValueOrDefault(stemElem) == brElem)           return "洩";
            if (LfElemOvercome.GetValueOrDefault(stemElem) == brElem)      return "財";
            return "比";
        }

        // 地支結構狀態："" | "被沖" | "被刑" | "合絆" | "合化XX"
        // isMonth=true 時月支才論化
        private static string LfRelBranchStructNote(string branch, string[] allBranches, bool isMonth, string mStem)
        {
            bool isChong = allBranches.Where(b => b != branch)
                .Any(b => LfChong.Contains(branch + b) || LfChong.Contains(b + branch));
            if (isChong) return "被沖";
            bool isXing = LfXing.Any(xg => xg.Contains(branch) && xg.Count(b => allBranches.Contains(b)) >= 2);
            if (isXing) return "被刑";
            if (LfHe.TryGetValue(branch, out var heInfo) && allBranches.Contains(heInfo.partner))
            {
                if (isMonth && KbStemToElement(mStem) == heInfo.elem) return $"合化{heInfo.elem}";
                return "合絆";
            }
            return "";
        }

        private static string LfBuildRelStemLine(
            string stem, string stemSS, string stemElem,
            string yongElem, string jiElem, string ageLabel, string domain)
        {
            string fav    = stemElem == yongElem ? "喜用" : stemElem == jiElem ? "忌神" : "閒神";
            bool isYin    = stemSS.Contains("印") || stemSS == "梟神";
            bool isCai    = stemSS.Contains("財");
            bool isGuan   = stemSS == "正官";
            bool isSha    = stemSS.Contains("七殺") || stemSS == "偏官";
            bool isShi    = stemSS == "食神";
            bool isShang  = stemSS == "傷官";
            bool isBi     = stemSS == "比肩";

            string desc;
            if (domain == "年")
            {
                if (fav == "喜用")
                {
                    if (isYin)   desc = "印星喜用，幼年出身書香有蔭，祖業有學識底蘊，長輩護持有力";
                    else if (isCai) desc = "財星喜用，幼年家境富足，父緣助力深，祖業財源充裕";
                    else if (isGuan) desc = "正官喜用，祖業有官聲規矩，出身有序，家教嚴謹有方向";
                    else if (isSha) desc = "七殺喜用，祖業剛毅有闖勁，幼年磨礪中成長，出身有韌性";
                    else if (isShi) desc = "食神喜用，祖業開明輕鬆，幼年生活豐足自在，家風和樂";
                    else if (isShang) desc = "傷官喜用，祖業才藝傳承，幼年聰慧活潑，家風重表現";
                    else if (isBi) desc = "比肩喜用，祖業同心協力，幼年家中自立精神強，出身積極";
                    else           desc = "劫財喜用，幼年同輩相助，家風重義氣，互助共進";
                }
                else if (fav == "忌神")
                {
                    if (isYin)   desc = "印星屬忌，幼年環境過保護或守舊，祖業守成難以拓展";
                    else if (isCai) desc = "財星屬忌，幼年家道財耗多，父緣受阻，家境起伏不穩";
                    else if (isGuan) desc = "正官屬忌，幼年規矩管束重，環境拘束壓抑，難以自由發展";
                    else if (isSha) desc = "七殺屬忌，幼年管教嚴苛，家庭壓力重，出身有波折磨難";
                    else if (isShi) desc = "食神屬忌，幼年散漫耗能，家境輕鬆但缺乏方向感與動力";
                    else if (isShang) desc = "傷官屬忌，幼年叛逆任性，祖業能量外洩，家風混亂";
                    else if (isBi) desc = "比肩屬忌，幼年資源被分散，家中自顧不暇，競爭多";
                    else           desc = "劫財屬忌，幼年家中財散，兄弟競爭多，祖業受損";
                }
                else
                    desc = "屬閒神，幼年出身中等，祖業平淡，家境無特殊助力";
            }
            else if (domain == "月")
            {
                if (fav == "喜用")
                {
                    if (isYin)   desc = "印綬喜用，父母重視教育，青年期學業有成，長輩庇蔭助力";
                    else if (isCai) desc = "財星喜用，父緣深厚，青年期財路漸開，家庭支援有力";
                    else if (isGuan) desc = "正官喜用，青年有目標方向，父母管教有方，功名仕途可期";
                    else if (isSha) desc = "七殺喜用，青年勇於挑戰，父母個性剛強但有助力，鍛鍊成材";
                    else if (isShi) desc = "食神喜用，青年才藝洋溢，父母開明鼓勵，求學輕鬆有成";
                    else if (isShang) desc = "傷官喜用，青年思維敏銳才藝出眾，父母鼓勵創意，表現突出";
                    else if (isBi) desc = "比肩喜用，青年自立心強，同儕相扶，奮發向上";
                    else           desc = "劫財喜用，青年義氣重，朋友助力多，同伴相扶共進";
                }
                else if (fav == "忌神")
                {
                    if (isYin)   desc = "印星屬忌，青年期依賴心重，求學動力不足，需自我督促";
                    else if (isCai) desc = "財星屬忌，父緣受阻，青年財路多波折，家庭支援有限";
                    else if (isGuan) desc = "正官屬忌，青年受規矩束縛，壓力沉重，發展受到限制";
                    else if (isSha) desc = "七殺屬忌，青年叛逆不喜約束，與父母常有摩擦，行事衝動";
                    else if (isShi) desc = "食神屬忌，青年散漫耗能，情緒起伏多，需收心聚焦";
                    else if (isShang) desc = "傷官屬忌，青年言語衝動易惹禍，父緣生疏，行事不拘";
                    else if (isBi) desc = "比肩屬忌，青年競爭損耗多，同儕摩擦大，需獨立打拚";
                    else           desc = "劫財屬忌，青年財散難靠兄弟，需謹慎理財方能守成";
                }
                else
                    desc = "屬閒神，青年父母緣分平淡，各自生活，需自立打拚";
            }
            else // 時柱
            {
                if (fav == "喜用")
                {
                    if (isYin)   desc = "印星喜用，晚年有靠山，子女孝順，學識涵養佳，老年生活安穩";
                    else if (isCai) desc = "財星喜用，晚年財源不斷，子女在財務上有助，老運豐足";
                    else if (isGuan) desc = "正官喜用，晚年受人尊重，子女有出息，社會地位持續提升";
                    else if (isSha) desc = "七殺喜用，晚年奮進不懈，子女個性強但有成就，老運有力";
                    else if (isShi) desc = "食神喜用，子女聰明孝順，晚年享子女之福，生活悠閒豐足";
                    else if (isShang) desc = "傷官喜用，子女才藝出眾，晚年因子女開心，享晚年福氣";
                    else if (isBi) desc = "比肩喜用，晚年友朋相扶，自立中有人相助，老運安康";
                    else           desc = "劫財喜用，晚年義氣深重，子女情誼濃，助力伴到老";
                }
                else if (fav == "忌神")
                {
                    if (isYin)   desc = "印星屬忌，晚年過分依賴，子女互動有距離，需主動溝通維繫";
                    else if (isCai) desc = "財星屬忌，晚年財耗多，子女緣薄，宜量入為出早做規劃";
                    else if (isGuan) desc = "正官屬忌，晚年仍受束縛，子女緣分一般，老運偏勞碌";
                    else if (isSha) desc = "七殺屬忌，晚年壓力仍重，子女有個性衝突，老運多奔波";
                    else if (isShi) desc = "食神屬忌，晚年精力外洩，子女緣分淡，老年需靜養保健";
                    else if (isShang) desc = "傷官屬忌，晚年言行易惹麻煩，子女難靠，老運多波折";
                    else if (isBi) desc = "比肩屬忌，晚年資源被分，子女難靠，需早謀晚年自立";
                    else           desc = "劫財屬忌，晚年財散子女難靠，需早做晚年財務規劃";
                }
                else
                    desc = "屬閒神，晚年子女緣分平淡，各自生活，老年需保持自立";
            }
            return $"{ageLabel}：{desc}。";
        }

        private static string LfBuildRelBranchLine(
            string branch, string branchSS, bool hasRoot, string effect, string stemFav,
            string structNote, string yongElem, string jiElem,
            string ageLabel, string domain, int gender)
        {
            bool isHua  = structNote.StartsWith("合化");
            bool isChong = structNote == "被沖";
            bool isXing  = structNote == "被刑";
            bool isHe    = structNote == "合絆";

            string text;
            if (isHua)
            {
                string huaElem = structNote.Replace("合化", "");
                string huaFav  = huaElem == yongElem ? "喜用" : huaElem == jiElem ? "忌神" : "閒神";
                text = huaFav == "喜用"
                    ? $"合化{huaElem}（喜用），合化後轉吉，此段緣分因合化而大為改善，助力增強"
                    : huaFav == "忌神"
                    ? $"合化{huaElem}（忌神），合化後轉凶，此段緣分因合化而受阻，阻力加重"
                    : $"合化{huaElem}，合化後化為閒神，此段緣分趨於平淡";
            }
            else if (isChong)
            {
                text = stemFav == "喜用"
                    ? "被沖動盪，喜用之力受破，此段緣分多起伏變動"
                    : "被沖制忌，忌神受破，阻力有所減輕但仍不安穩";
            }
            else if (isXing)
            {
                text = stemFav == "喜用"
                    ? "刑傷不安，喜用受損，此段緣分帶有磨折"
                    : "刑傷削忌，阻力雖減但有內傷，需謹慎應對";
            }
            else if (isHe)
            {
                text = stemFav == "喜用"
                    ? "合絆他處，喜用之力部分轉移，此段助力有所分散"
                    : "合絆制忌，忌神被牽制，阻力有所減輕";
            }
            else if (hasRoot)
            {
                text = (stemFav, domain) switch
                {
                    ("喜用", "年支") => "通根有力，家境環境穩固，助力貫穿整個童年",
                    ("喜用", "月支") => "通根有力，青壯期兄弟父母助力持續，環境條件有力支撐",
                    ("喜用", "配偶") => gender == 1
                        ? "通根有力，妻緣深厚，配偶賢良可靠，婚姻穩固助力，中年因婚而成"
                        : "通根有力，夫緣深厚，配偶有擔當，婚姻穩固助力，中年因婚而成",
                    ("喜用", "時支") => "通根有力，晚年子女緣分深厚，老年有依靠，根基穩固",
                    ("忌神", "年支") => "通根有力，忌神得根，家境阻力延續，整個童年環境持續受限",
                    ("忌神", "月支") => "通根有力，忌神得根，青壯期環境阻力延續，難以突破",
                    ("忌神", "配偶") => "通根有力，忌神有根，配偶緣分有阻，婚姻波折較多，需包容溝通",
                    ("忌神", "時支") => "通根有力，忌神有根，晚年子女緣薄，老年壓力持續",
                    _ => stemFav == "喜用" ? "通根有力，此段緣分深厚，助力穩固持續" : "通根有力，忌神有根，阻力延續難解"
                };
            }
            else
            {
                string domLabel = domain switch
                {
                    "年支" => "童年環境", "月支" => "青壯期",
                    "配偶" => "婚姻",    "時支" => "晚年", _ => "此段"
                };
                text = (effect, stemFav) switch
                {
                    ("生", "喜用") => $"無根得生，{domLabel}仍有一定助力，惟力度有限難以持久",
                    ("生", "忌神") => $"無根得生，忌神得助，{domLabel}阻礙延續",
                    ("剋", "喜用") => $"無根受剋，喜用被制，{domLabel}助力受阻，後段出現轉折",
                    ("剋", "忌神") => $"無根受剋，忌神被制，{domLabel}阻力有所緩解，可見轉機",
                    ("洩", "喜用") => $"無根被洩，喜用能量散失，{domLabel}助力逐漸消散",
                    ("洩", "忌神") => $"無根被洩，忌神消耗，{domLabel}阻力漸輕",
                    ("財", "喜用") => $"財氣加持，{domLabel}喜用得財相助，助益尚佳",
                    ("財", "忌神") => $"財反助忌，{domLabel}阻力因財而增，需謹慎",
                    _              => $"無根比和，{domLabel}助力平穩但力度有限"
                };
            }
            return $"{ageLabel}：{text}。";
        }

        private static string LfBuildCh5SixRelatives(
            string yStem, string yBranch, string mStem, string mBranch,
            string dStem, string dBranch, string hStem, string hBranch,
            string yStemSS, string mStemSS, string hStemSS,
            string yBranchSS, string mBranchSS, string dBranchSS, string hBranchSS,
            string yongShenElem, string jiShenElem, string dmElem,
            string bodyLabel, double bodyPct, int gender, string[] branches)
        {
            var sb = new StringBuilder();
            sb.AppendLine("【第五章：六親論斷】");

            // --- 年柱 ---
            sb.AppendLine("【年柱·出身幼年】");
            string yStemElem = KbStemToElement(yStem);
            string yStemFav  = yStemElem == yongShenElem ? "喜用" : yStemElem == jiShenElem ? "忌神" : "閒神";
            sb.AppendLine("  " + LfBuildRelStemLine(yStem, yStemSS, yStemElem, yongShenElem, jiShenElem, "幼年", "年"));
            sb.AppendLine("  " + LfBuildRelBranchLine(yBranch, yBranchSS,
                LfRelHasRoot(yStem, yBranch), LfRelBranchEffect(yBranch, yStem), yStemFav,
                LfRelBranchStructNote(yBranch, branches, false, mStem),
                yongShenElem, jiShenElem, "童年", "年支", gender));
            sb.AppendLine();

            // --- 月柱 ---
            sb.AppendLine("【月柱·父母青年】");
            string mStemElem = KbStemToElement(mStem);
            string mStemFav  = mStemElem == yongShenElem ? "喜用" : mStemElem == jiShenElem ? "忌神" : "閒神";
            sb.AppendLine("  " + LfBuildRelStemLine(mStem, mStemSS, mStemElem, yongShenElem, jiShenElem, "青年", "月"));
            sb.AppendLine("  " + LfBuildRelBranchLine(mBranch, mBranchSS,
                LfRelHasRoot(mStem, mBranch), LfRelBranchEffect(mBranch, mStem), mStemFav,
                LfRelBranchStructNote(mBranch, branches, true, mStem),
                yongShenElem, jiShenElem, "青壯", "月支", gender));
            sb.AppendLine();

            // --- 日柱 ---
            sb.AppendLine("【日柱·自身婚姻】");
            string selfDesc = bodyPct >= 55
                ? "自主能力足，事業主導力強，中年宜主動進取，可開創一番局面"
                : "需靠印比助身，宜藉助貴人環境之力，中年宜守成穩健為主";
            sb.AppendLine($"  中年：日主{bodyLabel}（{bodyPct:F0}%），{selfDesc}。");
            // 配偶宮獨看：以日支本氣十神及喜忌論配偶
            string dBranchElem = LfBranchHiddenRatio.TryGetValue(dBranch, out var dbh) && dbh.Count > 0
                ? KbStemToElement(dbh[0].stem) : "";
            string spouseStarElem = gender == 1
                ? LfElemOvercome.GetValueOrDefault(dmElem, "")
                : LfElemOvercomeBy.GetValueOrDefault(dmElem, "");
            bool dIsSpouseStar = dBranchElem == spouseStarElem;
            string dBranchFav  = dBranchElem == yongShenElem ? "喜用" : dBranchElem == jiShenElem ? "忌神" : "閒神";
            string dBranchLine = LfBuildRelBranchLine(dBranch, dBranchSS,
                LfRelHasRoot(dStem, dBranch), LfRelBranchEffect(dBranch, dStem), dBranchFav,
                LfRelBranchStructNote(dBranch, branches, false, mStem),
                yongShenElem, jiShenElem, "配偶", "配偶", gender);
            if (dIsSpouseStar)
            {
                string ssLabel = gender == 1 ? "妻星得位，" : "夫星得位，";
                dBranchLine = dBranchLine.Replace("配偶：", "配偶：" + ssLabel);
            }
            sb.AppendLine("  " + dBranchLine);
            sb.AppendLine();

            // --- 時柱 ---
            sb.AppendLine("【時柱·子女晚運】");
            string hStemElem = KbStemToElement(hStem);
            string hStemFav  = hStemElem == yongShenElem ? "喜用" : hStemElem == jiShenElem ? "忌神" : "閒神";
            sb.AppendLine("  " + LfBuildRelStemLine(hStem, hStemSS, hStemElem, yongShenElem, jiShenElem, "晚年", "時"));
            sb.AppendLine("  " + LfBuildRelBranchLine(hBranch, hBranchSS,
                LfRelHasRoot(hStem, hBranch), LfRelBranchEffect(hBranch, hStem), hStemFav,
                LfRelBranchStructNote(hBranch, branches, false, mStem),
                yongShenElem, jiShenElem, "老運", "時支", gender));

            return sb.ToString();
        }

        // 婚配屬相：依年支三合+六合找出有利婚配生肖
        private static string LfBuildSpouseZodiac(string yBranch)
        {
            var sanHe = new Dictionary<string, string[]>
            {
                {"申",new[]{"子","辰"}},{"子",new[]{"申","辰"}},{"辰",new[]{"申","子"}},
                {"亥",new[]{"卯","未"}},{"卯",new[]{"亥","未"}},{"未",new[]{"亥","卯"}},
                {"寅",new[]{"午","戌"}},{"午",new[]{"寅","戌"}},{"戌",new[]{"寅","午"}},
                {"巳",new[]{"酉","丑"}},{"酉",new[]{"巳","丑"}},{"丑",new[]{"巳","酉"}}
            };
            var liuHe = new Dictionary<string, string>
            {
                {"子","丑"},{"丑","子"},{"寅","亥"},{"亥","寅"},
                {"卯","戌"},{"戌","卯"},{"辰","酉"},{"酉","辰"},
                {"巳","申"},{"申","巳"},{"午","未"},{"未","午"}
            };
            var branchAnimal = new Dictionary<string, string>
            {
                {"子","鼠"},{"丑","牛"},{"寅","虎"},{"卯","兔"},{"辰","龍"},{"巳","蛇"},
                {"午","馬"},{"未","羊"},{"申","猴"},{"酉","雞"},{"戌","狗"},{"亥","豬"}
            };
            var compatible = new List<string>();
            if (sanHe.TryGetValue(yBranch, out var sh))
                compatible.AddRange(sh.Select(b => branchAnimal.GetValueOrDefault(b, b)));
            if (liuHe.TryGetValue(yBranch, out var lh))
                compatible.Add(branchAnimal.GetValueOrDefault(lh, lh));
            if (compatible.Count == 0) return "";
            return $"婚配屬相：{string.Join("、", compatible)} 相合，對感情與婚姻最有助益。";
        }

        // 五行人事物類象說明（第十章用）
        private static string LfBuildWuXingLeiXiang(string yongElem, string jiElem)
        {
            var elemPeople = new Dictionary<string, string>
            {
                {"木","文人學者、老師、設計師、律師、醫生"},
                {"火","藝術家、演講者、科技網路人才、廚師"},
                {"土","建商、地產從業者、農業、行政人員"},
                {"金","銀行業者、軍警、工程師、外科醫生"},
                {"水","業務貿易、旅遊導覽、媒體傳播、流通業者"}
            };
            var elemThings = new Dictionary<string, string>
            {
                {"木","書籍文具、植物木材、文創品、東方位置"},
                {"火","電子設備、燈光熱食、表演場所、南方位置"},
                {"土","土地房產、磚石陶器、農產品、中央位置"},
                {"金","金屬珠寶、機械設備、契約文件、西方位置"},
                {"水","飲料水源、流通物品、水域場所、北方位置"}
            };
            var elemEvents = new Dictionary<string, string>
            {
                {"木","文書簽約、學習進修、爭取名位、貴人引薦"},
                {"火","名聲曝光、人際往來、表演發表、慶典儀式"},
                {"土","置產投資、穩健守成、信任合作、長期佈局"},
                {"金","財務清算、合約執行、競爭考試、官方往來"},
                {"水","旅行移動、資金流轉、溝通談判、靈活應變"}
            };
            var sb = new StringBuilder();
            sb.AppendLine($"用神（{yongElem}）人事物類象：");
            sb.AppendLine($"  宜親近：{elemPeople.GetValueOrDefault(yongElem, "")}");
            sb.AppendLine($"  有利事物：{elemThings.GetValueOrDefault(yongElem, "")}");
            sb.AppendLine($"  有利時機：{elemEvents.GetValueOrDefault(yongElem, "")}");
            sb.AppendLine($"忌神（{jiElem}）人事物類象（謹慎應對）：");
            sb.AppendLine($"  宜謹慎：{elemPeople.GetValueOrDefault(jiElem, "")}");
            sb.AppendLine($"  不利事物：{elemThings.GetValueOrDefault(jiElem, "")}");
            sb.Append($"  需留意：{elemEvents.GetValueOrDefault(jiElem, "")}");
            return sb.ToString();
        }

        // 健康注意：天干被剋及地支刑沖提醒（第九章用）
        private static string LfBuildHealthWarnings(
            string yStem, string mStem, string dStem, string hStem,
            string yBranch, string mBranch, string dBranch, string hBranch)
        {
            var warnings = new List<string>();
            var stemOrgan = new Dictionary<string, string>
            {
                {"甲","肝膽"},{"乙","肝膽"},{"丙","心血管"},{"丁","心血管"},
                {"戊","脾胃腸"},{"己","脾胃腸"},{"庚","肺呼吸"},{"辛","肺呼吸"},
                {"壬","腎泌尿"},{"癸","腎泌尿"}
            };
            var pillarNames = new[] { "年", "月", "日", "時" };
            var pillarBodyStem = new[] { "頭頸", "胸肩", "腰腹", "腿腳" };
            var pillarStem = new[] { yStem, mStem, dStem, hStem };
            var allBranches = new[] { yBranch, mBranch, dBranch, hBranch };
            var allStemElems = pillarStem.Select(KbStemToElement).ToArray();
            // 天干合（合剋不顯示）
            var tianGanHe = new HashSet<string> { "甲己","己甲","乙庚","庚乙","丙辛","辛丙","丁壬","壬丁","戊癸","癸戊" };
            // 天干被剋：只查相鄰柱（年月、月日、日時），且排除天干合
            var adjacentPairs = new[] { (0, 1), (1, 2), (2, 3) };
            foreach (var (i, j) in adjacentPairs)
            {
                // 檢查 j 的元素是否剋 i（即 i 被 j 攻）
                string elemI = allStemElems[i];
                string elemJ = allStemElems[j];
                string overcomeByI = LfElemOvercomeBy.GetValueOrDefault(elemI, "");
                bool iAttackedByJ = elemJ == overcomeByI;
                // 反向：i 剋 j
                string overcomeByJ = LfElemOvercomeBy.GetValueOrDefault(elemJ, "");
                bool jAttackedByI = elemI == overcomeByJ;

                if (iAttackedByJ && !tianGanHe.Contains(pillarStem[i] + pillarStem[j]))
                {
                    string organ = stemOrgan.GetValueOrDefault(pillarStem[i], "");
                    warnings.Add($"{pillarNames[i]}柱天干受剋，{organ}（{pillarBodyStem[i]}）較為脆弱，宜注意{organ}保健。");
                }
                if (jAttackedByI && !tianGanHe.Contains(pillarStem[j] + pillarStem[i]))
                {
                    string organ = stemOrgan.GetValueOrDefault(pillarStem[j], "");
                    warnings.Add($"{pillarNames[j]}柱天干受剋，{organ}（{pillarBodyStem[j]}）較為脆弱，宜注意{organ}保健。");
                }
            }
            // 地支六沖
            var chongPairs = new Dictionary<string, string>
            {
                {"子","午"},{"午","子"},{"丑","未"},{"未","丑"},
                {"寅","申"},{"申","寅"},{"卯","酉"},{"酉","卯"},
                {"辰","戌"},{"戌","辰"},{"巳","亥"},{"亥","巳"}
            };
            var branchOrgan = new Dictionary<string, string>
            {
                {"子","腎泌尿"},{"亥","腎泌尿"},{"丑","脾胃"},{"辰","脾胃"},{"未","脾胃"},{"戌","脾胃"},
                {"寅","肝膽"},{"卯","肝膽"},{"巳","心血管"},{"午","心血管"},{"申","肺呼吸"},{"酉","肺呼吸"}
            };
            for (int i = 0; i < 4; i++)
            {
                if (!chongPairs.TryGetValue(allBranches[i], out string? chongWith)) continue;
                int j = Array.IndexOf(allBranches, chongWith, i + 1);
                if (j < 0) continue;
                string organ = branchOrgan.GetValueOrDefault(allBranches[i], "");
                warnings.Add($"{pillarNames[i]}支{allBranches[i]}與{pillarNames[j]}支{allBranches[j]}相沖，{organ}系統宜留意，中年後建議定期檢查。");
            }
            // 地支三刑
            var xingGroups = new[] {
                new[] { "寅", "巳", "申" },
                new[] { "丑", "戌", "未" },
                new[] { "子", "卯" }
            };
            foreach (var xg in xingGroups)
            {
                var matched = xg.Where(b => allBranches.Contains(b)).ToArray();
                if (matched.Length >= 2)
                {
                    var xingOrgans = matched.Select(b => branchOrgan.GetValueOrDefault(b, ""))
                        .Where(o => !string.IsNullOrEmpty(o)).Distinct().ToList();
                    if (xingOrgans.Count > 0)
                        warnings.Add($"命局帶{string.Join("、", matched)}刑，{string.Join("、", xingOrgans)}系統需注意保養，情緒管理亦需留意。");
                }
            }
            if (warnings.Count == 0) return "";
            return string.Join("\n", warnings.Select(w => $"  ★ {w}"));
        }

        // 建立天干地支喜忌對照表（第四章用，pipe table 格式）
        private static string LfBuildYongJiTable(
            string yongShenElem, string fuYiElem, string jiShenElem, string tuneElem,
            string dStem, string[] chartBranches)
        {
            string jiYongElem = LfElemOvercomeBy.GetValueOrDefault(yongShenElem, "");

            string Cls(string elem)
            {
                if (elem == jiShenElem) return "X";
                if (elem == yongShenElem || elem == fuYiElem ||
                    (!string.IsNullOrEmpty(tuneElem) && elem == tuneElem)) return "○";
                if (elem == jiYongElem && jiYongElem != jiShenElem) return "△忌";
                return "△";
            }

            string[] stems = { "甲","乙","丙","丁","戊","己","庚","辛","壬","癸" };
            string[] brs   = { "子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥" };
            string sepS  = "|:---:|" + string.Join("|", Enumerable.Repeat(":---:", stems.Length)) + "|";
            string sepB  = "|:---:|" + string.Join("|", Enumerable.Repeat(":---:", brs.Length)) + "|";

            var sb = new StringBuilder();

            // 天干喜忌表格（橫向：天干當欄，三列屬性）
            sb.AppendLine("【天干喜忌對照】");
            sb.AppendLine("| 天干 | " + string.Join(" | ", stems) + " |");
            sb.AppendLine("|:---:|" + string.Join("|", stems.Select(_ => ":---:")) + "|");
            sb.AppendLine("| 五行 | " + string.Join(" | ", stems.Select(s => KbStemToElement(s))) + " |");
            sb.AppendLine("| 十神 | " + string.Join(" | ", stems.Select(s => LfStemShiShen(s, dStem))) + " |");
            sb.AppendLine("| 喜忌 | " + string.Join(" | ", stems.Select(s => Cls(KbStemToElement(s)))) + " |");
            sb.AppendLine();

            // 地支喜忌表格（橫向：地支當欄，三列屬性，命局地支加*）
            sb.AppendLine("【地支喜忌對照】（*=命局已有）");
            sb.AppendLine("| 地支 | " + string.Join(" | ", brs.Select(br => chartBranches.Contains(br) ? br + "*" : br)) + " |");
            sb.AppendLine("|:---:|" + string.Join("|", brs.Select(_ => ":---:")) + "|");
            sb.AppendLine("| 五行 | " + string.Join(" | ", brs.Select(br =>
                LfBranchHiddenRatio.TryGetValue(br, out var h) && h.Count > 0 ? KbStemToElement(h[0].stem) : "-")) + " |");
            sb.AppendLine("| 十神 | " + string.Join(" | ", brs.Select(br => {
                string ms = LfBranchHiddenRatio.TryGetValue(br, out var h2) && h2.Count > 0 ? h2[0].stem : "";
                return !string.IsNullOrEmpty(ms) ? LfStemShiShen(ms, dStem) : "-";
            })) + " |");
            sb.AppendLine("| 喜忌 | " + string.Join(" | ", brs.Select(br => {
                string elem = LfBranchHiddenRatio.TryGetValue(br, out var h3) && h3.Count > 0
                    ? KbStemToElement(h3[0].stem) : "";
                return elem != "" ? Cls(elem) : "-";
            })) + " |");

            sb.AppendLine("說明：○=喜用  △忌=次忌（克用神）  △=中性  X=大忌（克身）");
            return sb.ToString();
        }

        // === Lf Report Builder ===

        // ─── 小人類型分析 ───────────────────────────────────────────────────
        private static string LfXiaoRenAnalysis(
            string yStem, string yBranch,
            string mStem, string mBranch,
            string dStem, string dBranch,
            string hStem, string hBranch,
            string jiShenElem, string dmElem)
        {
            var sb = new StringBuilder();
            string jiSsName;
            if (jiShenElem == LfElemOvercomeBy.GetValueOrDefault(dmElem, ""))
                jiSsName = "官殺性質（上司、公權力、強勢競爭者）";
            else if (jiShenElem == dmElem)
                jiSsName = "比劫性質（同業競爭者、利益衝突同階層）";
            else if (jiShenElem == LfElemOvercome.GetValueOrDefault(dmElem, ""))
                jiSsName = "財星性質（財務往來關係人、利益衝突者）";
            else if (jiShenElem == LfGenByElem.GetValueOrDefault(dmElem, ""))
                jiSsName = "印梟性質（長輩顧問型、表面關心者）";
            else if (jiShenElem == LfElemGen.GetValueOrDefault(dmElem, ""))
                jiSsName = "食傷性質（部屬晚輩型、口舌是非者）";
            else
                jiSsName = "";

            sb.Append($"忌神五行【{jiShenElem}】");
            if (!string.IsNullOrEmpty(jiSsName)) sb.AppendLine($"（{jiSsName}）");
            else sb.AppendLine();
            sb.AppendLine("忌神在四柱出現位置（決定小人身份方向）：");

            var pillars = new[]
            {
                (stem: yStem, branch: yBranch, label: "年柱", who: "長輩、父母、上一輩"),
                (stem: mStem, branch: mBranch, label: "月柱", who: "同輩、同事、同學、兄弟姐妹"),
                (stem: dStem, branch: dBranch, label: "日柱", who: "自身判斷易偏差，需防伴侶"),
                (stem: hStem, branch: hBranch, label: "時柱", who: "晚輩、子女、部屬"),
            };
            bool anyFound = false;
            var jiZodiacs = new List<string>();
            foreach (var (stem, branch, label, who) in pillars)
            {
                string stemElem = KbStemToElement(stem);
                string branchMainElem = LfBranchElem.GetValueOrDefault(branch, "");
                bool stemIsJi = stemElem == jiShenElem;
                bool branchIsJi = branchMainElem == jiShenElem;
                if (stemIsJi || branchIsJi)
                {
                    anyFound = true;
                    sb.AppendLine($"  {label}見忌神 → 防範來自【{who}】的小人");
                    if (branchIsJi && LfBranchZodiac.TryGetValue(branch, out var z))
                        jiZodiacs.Add(z);
                }
            }
            if (!anyFound)
                sb.AppendLine("  四柱中忌神分布較輕，小人傷害相對有限。");

            if (jiZodiacs.Count > 0)
                sb.AppendLine($"忌神生肖（深交需謹慎）：{string.Join("、", jiZodiacs.Distinct())}");

            string chongBranch = LfBranchChongOf.GetValueOrDefault(yBranch, "");
            if (!string.IsNullOrEmpty(chongBranch) && LfBranchZodiac.TryGetValue(chongBranch, out var chongZ))
                sb.AppendLine($"事業朋友圈謹慎往來：{chongZ}生（與本命年支{yBranch}相沖，利益摩擦機率高）");

            return sb.ToString().TrimEnd();
        }

        // ─── 官司文書風險分析 ─────────────────────────────────────────────
        private static string LfGuanSiAnalysis(
            string yStem, string yBranch,
            string mStem, string mBranch,
            string dStem, string dBranch,
            string hStem, string hBranch,
            string jiShenElem, string dmElem, double bodyPct)
        {
            var sb = new StringBuilder();
            string guanShaElem = LfElemOvercomeBy.GetValueOrDefault(dmElem, "");
            string yinXiaoElem = LfGenByElem.GetValueOrDefault(dmElem, "");
            string shishangElem = LfElemGen.GetValueOrDefault(dmElem, "");
            string caiElem = LfElemOvercome.GetValueOrDefault(dmElem, "");
            bool isWeak = bodyPct < 50;
            string[] allStems = { yStem, mStem, dStem, hStem };
            string[] allBranches = { yBranch, mBranch, dBranch, hBranch };
            bool hasShiShang = allStems.Any(s => KbStemToElement(s) == shishangElem)
                || allBranches.Any(b => LfBranchElem.GetValueOrDefault(b, "") == shishangElem);
            bool hasCai = allStems.Any(s => KbStemToElement(s) == caiElem)
                || allBranches.Any(b => LfBranchElem.GetValueOrDefault(b, "") == caiElem);

            if (isWeak)
            {
                sb.AppendLine($"日主身弱，官司風險來自忌神官殺（{guanShaElem}）。");
                if (jiShenElem == guanShaElem)
                {
                    if (hasShiShang)
                        sb.AppendLine($"  八字有食傷（{shishangElem}）制官殺 → 即使遇官司糾紛，有化解空間，問題不大。");
                    else
                        sb.AppendLine($"  八字無食傷制官殺 → 官司一旦發生，影響較為嚴重，需謹慎處理、及早和解。");
                }
                else
                    sb.AppendLine($"  忌神並非官殺，先天官司風險相對較低，但仍需注意公權力衝突。");
            }
            else
            {
                sb.AppendLine($"日主身強，文書名譽風險來自忌神印梟（{yinXiaoElem}）。");
                if (jiShenElem == yinXiaoElem)
                {
                    if (hasCai)
                        sb.AppendLine($"  八字有財（{caiElem}）制印梟 → 文書糾紛雖可能發生，財力或人際可化解。");
                    else
                        sb.AppendLine($"  八字無財制印梟 → 文書名譽受損一旦發生，影響難以收拾，合約與名聲務必謹慎。");
                }
                else
                    sb.AppendLine($"  忌神並非印梟，先天文書官司風險相對較低。");
            }

            var pillars = new[]
            {
                (stem: yStem, branch: yBranch, label: "年柱", who: "長輩或父母"),
                (stem: mStem, branch: mBranch, label: "月柱", who: "同輩、同事或合夥人"),
                (stem: dStem, branch: dBranch, label: "日柱", who: "自身或伴侶"),
                (stem: hStem, branch: hBranch, label: "時柱", who: "晚輩或子女"),
            };
            string triggerElem = isWeak ? guanShaElem : yinXiaoElem;
            bool foundPillar = false;
            foreach (var (stem, branch, label, who) in pillars)
            {
                string stemElem = KbStemToElement(stem);
                string branchMainElem = LfBranchElem.GetValueOrDefault(branch, "");
                if (stemElem == triggerElem || branchMainElem == triggerElem)
                {
                    if (!foundPillar) { sb.AppendLine("涉事對象方向："); foundPillar = true; }
                    sb.AppendLine($"  {label}見起因星 → 糾紛可能來自【{who}】");
                }
            }
            return sb.ToString().TrimEnd();
        }

        // ─── 車關時機分析 ──────────────────────────────────────────────────
        private static string LfCheGuanAnalysis(
            string yBranch, string mBranch,
            string dBranch, string hBranch,
            string jiShenElem, string dmElem)
        {
            var sb = new StringBuilder();
            string guanShaElem = LfElemOvercomeBy.GetValueOrDefault(dmElem, "");
            var siYiMa = new HashSet<string> { "寅", "申", "巳", "亥" };
            var riskBranches = new List<(string branch, string label)>();
            foreach (var (branch, label) in new[] {(yBranch,"年支"),(mBranch,"月支"),(dBranch,"日支"),(hBranch,"時支")})
            {
                string branchElem = LfBranchElem.GetValueOrDefault(branch, "");
                if (siYiMa.Contains(branch) && branchElem == guanShaElem)
                    riskBranches.Add((branch, label));
            }

            if (riskBranches.Count == 0)
            {
                sb.AppendLine("四柱驛馬位（寅申巳亥）無官殺，先天車關風險相對較低。");
            }
            else
            {
                sb.AppendLine("先天八字驛馬含官殺，有車關傾向：");
                foreach (var (branch, label) in riskBranches)
                    sb.AppendLine($"  {label}（{branch}）為{guanShaElem}官殺驛馬 → 此地支被沖合之年需特別注意");
                sb.AppendLine("引動時機：大運或流年地支與上述驛馬官殺形成六沖、三合、三會之年，車關風險提升。");
                sb.AppendLine("建議：行凶運期間避免長途自駕，注意交通安全，盡量搭乘大眾運輸。");
            }
            return sb.ToString().TrimEnd();
        }

        // ─── 海外發展分析 ─────────────────────────────────────────────────
        private static string LfHaiWaiAnalysis(
            string yBranch, string mBranch,
            string dBranch, string hBranch,
            string yongShenElem, string jiShenElem, string dmElem,
            bool hasZiwei, JsonElement palaces)
        {
            var sb = new StringBuilder();
            var siYiMa = new HashSet<string> { "寅", "申", "巳", "亥" };
            string guanShaElem = LfElemOvercomeBy.GetValueOrDefault(dmElem, "");
            string yinXiaoElem = LfGenByElem.GetValueOrDefault(dmElem, "");
            bool yearMaIsYong = siYiMa.Contains(yBranch)
                && LfBranchElem.GetValueOrDefault(yBranch, "") == yongShenElem;
            bool monthMaIsYong = siYiMa.Contains(mBranch)
                && LfBranchElem.GetValueOrDefault(mBranch, "") == yongShenElem;
            bool hasBaziSignal = yearMaIsYong || monthMaIsYong;

            if (hasBaziSignal)
            {
                sb.AppendLine("八字先天出行信號：年月柱有喜用神驛馬（寅申巳亥），出國機會有利。");
                if (yongShenElem == guanShaElem)
                    sb.AppendLine("  → 官殺為喜用：適合赴海外工作、外派或接受官方職務。");
                else if (yongShenElem == yinXiaoElem)
                    sb.AppendLine("  → 印梟為喜用：適合海外長期定居、移居或求學。");
                else
                    sb.AppendLine("  → 出國有利，可把握海外發展機會。");
            }
            else
                sb.AppendLine("八字年月柱驛馬不具喜用五行，先天出國發展優勢較不明顯。");

            if (hasZiwei && palaces.ValueKind == JsonValueKind.Array)
            {
                try
                {
                    foreach (var p in palaces.EnumerateArray())
                    {
                        string palName = p.TryGetProperty("name", out var pn) ? pn.GetString() ?? "" : "";
                        string palBranch = p.TryGetProperty("branch", out var pb) ? pb.GetString() ?? "" : "";
                        if (!siYiMa.Contains(palBranch)) continue;
                        bool hasSiHuaGood = false;
                        if (p.TryGetProperty("stars", out var stars) && stars.ValueKind == JsonValueKind.Array)
                            hasSiHuaGood = stars.EnumerateArray().Any(s =>
                            {
                                string sn = s.TryGetProperty("name", out var n) ? n.GetString() ?? "" : "";
                                return sn.Contains("化祿") || sn.Contains("化科");
                            });
                        if (!hasSiHuaGood) continue;
                        string palLabel = palName.TrimEnd('宮');
                        if (palLabel == "遷移")
                            sb.AppendLine($"紫微遷移宮（{palBranch}）有化祿/化科，位於驛馬宮位 → 出國運勢強旺，宜把握海外機遇。");
                        else if (palLabel == "田宅")
                            sb.AppendLine($"紫微田宅宮（{palBranch}）有化祿/化科，位於驛馬宮位 → 有移居定居海外之象。");
                    }
                }
                catch { /* 紫微解析異常時略過 */ }
            }

            if (!hasBaziSignal)
                sb.AppendLine("建議：可把握大運/流年驛馬引動之年短期出行或探索海外機會。");

            return sb.ToString().TrimEnd();
        }

        private static string LfBuildReport(
            string yStem, string yBranch, string mStem, string mBranch,
            string dStem, string dBranch, string hStem, string hBranch,
            string yStemSS, string mStemSS, string hStemSS,
            string yBranchSS, string mBranchSS, string dBranchSS, string hBranchSS,
            string dmElem, Dictionary<string, double> wuXing, double bodyPct, string bodyLabel,
            string season, string seaLabel, string pattern, string yongShenElem, string fuYiElem, string yongReason,
            string jiShenElem,
            List<(string stem, string branch, string liuShen, int startAge, int endAge, int score, string level)> scored,
            int gender, int birthYear)
        {
            var sb = new StringBuilder();
            string genderText = gender == 1 ? "男（乾造）" : "女（坤造）";
            var branches = new[] { yBranch, mBranch, dBranch, hBranch };
            string SS(string ss) => string.IsNullOrEmpty(ss) ? "" : $"（{ss}）";
            string LfWxSS(string e) => $"{e}{wuXing[e]:F0}%({LfElemSsGroup(e, dmElem)})";
            string wx = $"{LfWxSS("木")} {LfWxSS("火")} {LfWxSS("土")} {LfWxSS("金")} {LfWxSS("水")}";
            int currentAge = DateTime.Today.Year - birthYear;

            sb.AppendLine("=================================================================");
            sb.AppendLine("                         八 字 命 書");
            sb.AppendLine("=================================================================");
            sb.AppendLine();

            // 人生指南目錄
            sb.AppendLine("                       人  生  指  南");
            sb.AppendLine("-----------------------------------------------------------------");
            sb.AppendLine("  命盤基本資訊");
            sb.AppendLine("  命局體性（寒暖濕燥）");
            sb.AppendLine("  日主強弱判定");
            sb.AppendLine("  格局與用神判定");
            sb.AppendLine("  六親論斷");
            sb.AppendLine("  性格志向");
            sb.AppendLine("  事業財運");
            sb.AppendLine("  婚姻感情");
            sb.AppendLine("  健康壽元");
            sb.AppendLine("  一生命運總評");
            sb.AppendLine("  人生警示事項");
            sb.AppendLine("  適合行業建議");
            sb.AppendLine("  居家風水開運");
            sb.AppendLine("-----------------------------------------------------------------");
            sb.AppendLine();

            // === Ch.1 命盤基本資訊 ===
            sb.AppendLine("【第一章：命盤基本資訊】");
            sb.AppendLine($"性別：{genderText}  出生年：{birthYear} 年");
            sb.AppendLine($"四柱：{yStem}{yBranch} {mStem}{mBranch} {dStem}{dBranch} {hStem}{hBranch}");
            sb.AppendLine($"十神：年干{SS(yStemSS)} 年支{SS(yBranchSS)} 月干{SS(mStemSS)} 月支{SS(mBranchSS)} 時干{SS(hStemSS)} 時支{SS(hBranchSS)}");
            sb.AppendLine($"日主：{dStem}（{dmElem}）");
            if (scored.Count > 0)
                sb.AppendLine($"大運起運：{scored[0].startAge} 歲，依虛歲生日後換運為主");
            sb.AppendLine();

            // === Ch.2 命局體性 ===
            sb.AppendLine("【第二章：命局體性（寒暖濕燥）】");
            sb.AppendLine($"月支 {mBranch} 生人，命局屬【{seaLabel}】。");
            if (seaLabel == "寒凍")
                sb.AppendLine("最喜：丙丁巳午火暖局。最忌：壬癸亥子水助寒。調候急用：丙丁火。");
            else if (seaLabel == "炎熱")
                sb.AppendLine("最喜：壬癸亥子水消暑，庚辛申酉金。最忌：丙丁巳午火助熱。調候急用：壬癸水。");
            else
                sb.AppendLine("體性溫和，以日主強弱論用神，無需特別調候。");
            sb.AppendLine();

            // === Ch.3 日主強弱 ===
            sb.AppendLine("【第三章：日主強弱判定】");
            sb.AppendLine($"日干 {dStem}（{dmElem}），月令 {mBranch}（{season}季）。");
            sb.AppendLine($"五行分布：{wx}");
            double biJiPct = wuXing.GetValueOrDefault(dmElem, 0) + wuXing.GetValueOrDefault(LfGenByElem.GetValueOrDefault(dmElem, ""), 0);
            double tiPct = wuXing.GetValueOrDefault(dmElem, 0) + wuXing.GetValueOrDefault(LfGenByElem.GetValueOrDefault(dmElem, ""), 0) + wuXing.GetValueOrDefault(LfElemGen.GetValueOrDefault(dmElem, ""), 0);
            double yongPct = wuXing.GetValueOrDefault(LfElemOvercome.GetValueOrDefault(dmElem, ""), 0) + wuXing.GetValueOrDefault(LfElemOvercomeBy.GetValueOrDefault(dmElem, ""), 0);
            sb.AppendLine($"比印陣：{biJiPct:F0}% | 洩克陣：{100 - biJiPct:F0}%   (印比食)體 {tiPct:F0}%  (財官)用 {yongPct:F0}%");
            sb.AppendLine($"結論：日主【{bodyLabel}】（強弱度：{bodyPct:F0}%）");
            sb.AppendLine();

            // === Ch.4 格局與用神 ===
            sb.AppendLine("【第四章：格局與用神判定】");
            sb.AppendLine($"格局：【{pattern}】");
            sb.AppendLine($"用神：【{yongShenElem}】（理由：{yongReason}）");
            sb.AppendLine($"喜用：天干 {LfElemStems(yongShenElem)}，地支 {LfElemBranches(yongShenElem)}");
            if (fuYiElem != yongShenElem)
                sb.AppendLine($"輔助喜神：【{fuYiElem}】（{(bodyPct <= 40 ? "印比互補扶身" : "官財互補制衡")}）");
            string tuneElemDisp = season == "冬" ? "火" : season == "夏" ? "水" : "";
            if (!string.IsNullOrEmpty(tuneElemDisp) && tuneElemDisp != yongShenElem && tuneElemDisp != fuYiElem)
                sb.AppendLine($"調候喜神：【{tuneElemDisp}】（{(season == "冬" ? "冬月寒凍，喜火暖局" : "夏月炎熱，喜水消暑")}）");
            string jiYongElemDisp = LfElemOvercomeBy.GetValueOrDefault(yongShenElem, "");
            sb.AppendLine($"大忌(X)：{jiShenElem}，天干 {LfElemStems(jiShenElem)}，地支 {LfElemBranches(jiShenElem)}");
            if (!string.IsNullOrEmpty(jiYongElemDisp) && jiYongElemDisp != jiShenElem)
                sb.AppendLine($"次忌(△忌)：{jiYongElemDisp}，天干 {LfElemStems(jiYongElemDisp)}，地支 {LfElemBranches(jiYongElemDisp)}（克用神{yongShenElem}，力道較輕）");
            sb.AppendLine($"格局說明：{LfPatternDesc(pattern)}");
            sb.AppendLine();
            sb.Append(LfBuildYongJiTable(yongShenElem, fuYiElem, jiShenElem, tuneElemDisp, dStem, branches));

            // === Ch.5 六親論斷 ===
            sb.Append(LfBuildCh5SixRelatives(
                yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch,
                yStemSS, mStemSS, hStemSS, yBranchSS, mBranchSS, dBranchSS, hBranchSS,
                yongShenElem, jiShenElem, dmElem, bodyLabel, bodyPct, gender, branches));
            string rules = LfApplyRules(yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch,
                yStemSS, mStemSS, hStemSS, yBranchSS, mBranchSS, dBranchSS, hBranchSS,
                dmElem, wuXing, gender, pattern, bodyPct, branches);
            // 移除規則代碼（如 F01、B04、S01 等），保留純中文標籤
            rules = System.Text.RegularExpressions.Regex.Replace(rules, @"【[A-Z]\d+[^】\s]*\s+", "【");
            if (!string.IsNullOrEmpty(rules)) sb.AppendLine(rules);
            sb.AppendLine();

            // === Ch.6 性格志向 ===
            sb.AppendLine("【第六章：性格志向】");
            sb.AppendLine($"日主 {dStem}（{dmElem}），格局為{pattern}，日主{bodyLabel}（{bodyPct:F0}%）。");
            sb.AppendLine(LfPersonalityDesc(dmElem, pattern, bodyPct));
            sb.AppendLine();

            // === Ch.7 事業財運 ===
            sb.AppendLine("【第七章：事業財運】");
            double caiPct = wuXing.GetValueOrDefault(LfElemOvercome.GetValueOrDefault(dmElem, ""), 0);
            // 職業天份
            string careerGiftDesc = pattern switch
            {
                "正官格" => "正官格之人擁有天生的管理氣質，做事有條理、重規矩，在有制度的組織環境中最能出頭。適合公務機關、大型企業、法律、金融等需要信任與規範的行業。升遷往往靠口碑與資歷，穩扎穩打是您的致勝之道。",
                "七殺格" => "七殺格之人有強烈的鬥志與執行力，不怕困難，敢於挑戰高目標。適合需要競爭、對抗、衝勁的行業，如軍警、業務、外科醫療、競技運動等。也適合自行創業，但宜有紀律有計劃，避免衝動行事。",
                "食神格" => "食神格之人有創意與藝術氣質，工作上重視享受過程，適合餐飲、美食、設計、藝術、教育、休閒娛樂等行業。財運方面福星高照，不需過度強求，往往在喜歡的事情上自然得財。",
                "傷官格" => "傷官格之人才華突出，思維跳脫框架，適合創意、技術、表演、寫作、科技研發等需要腦力與創新的工作。在自由度高的環境中發揮最佳，不適合死板的上下班制度，獨立接案或自行創業更能展現才華。",
                "正財格" => "正財格之人腳踏實地，適合商業、財務、會計、零售、房地產等穩定收入的行業。財富靠努力積累，一分耕耘一分收穫，最忌投機取巧；以實力與信譽建立事業，是長遠之道。",
                "偏財格" => "偏財格之人人緣廣、善交際，適合業務、貿易、投資理財、公關、娛樂等需要廣結人脈的行業。財富來源多元，偏財機遇多，但需注意守財，避免財來財去。",
                "正印格" => "正印格之人學識淵博，思維嚴謹，適合學術研究、教育、文化出版、醫療護理、公共行政等重視專業素養的行業。貴人緣強，常能在關鍵時刻獲得提攜；靠學識與名聲立足，比靠金錢更能走得長遠。",
                "偏印格" => "偏印格之人思維獨特，適合宗教、哲學、神秘學、藝術、特殊技能、研究等偏門但深入的領域。不適合流水線式的普通工作，需要有能讓自己深度鑽研的舞台，才能真正發光發熱。",
                _ => $"格局為{pattern}，天生適合{LfCareerDesc(pattern)}的方向，用神五行所代表的行業是最佳選擇。"
            };
            sb.AppendLine(careerGiftDesc);
            sb.AppendLine();
            // 財富特質
            sb.AppendLine(LfWealthDesc(caiPct, bodyPct, wuXing.GetValueOrDefault(dmElem, 0)));
            sb.AppendLine();
            // 事業開運方向
            sb.AppendLine($"開運行業：{LfElemCareer(yongShenElem)}");
            sb.AppendLine($"開運方位：{LfElemDir(yongShenElem)}  開運色彩：{LfElemColor(yongShenElem)}");
            sb.AppendLine();

            // === Ch.8 婚姻感情 ===
            if (!LfShouldSkipPalace("夫妻宮", currentAge))
            {
                sb.AppendLine("【第八章：婚姻感情】");
                string spouseElem = gender == 1 ? LfElemOvercome.GetValueOrDefault(dmElem,"") : LfElemOvercomeBy.GetValueOrDefault(dmElem,"");
                double spousePct  = wuXing.GetValueOrDefault(spouseElem, 0);
                string spouseStar = gender == 1 ? "妻星（財）" : "夫星（官殺）";
                sb.AppendLine($"{spouseStar}五行：{spouseElem}，占命局 {spousePct:F0}%。");
                sb.AppendLine(LfMarriageDesc(spousePct, dBranch, dStem, dmElem, gender, branches));
                string spouseZodiac = LfBuildSpouseZodiac(yBranch);
                if (!string.IsNullOrEmpty(spouseZodiac)) sb.AppendLine(spouseZodiac);
                sb.AppendLine();
            }

            // === Ch.9 健康壽元 ===
            sb.AppendLine("【第九章：健康壽元】");
            sb.AppendLine($"命局五行：{wx}");
            sb.AppendLine(LfHealthDesc(wuXing, seaLabel));
            string healthWarnings = LfBuildHealthWarnings(yStem, mStem, dStem, hStem, yBranch, mBranch, dBranch, hBranch);
            if (!string.IsNullOrEmpty(healthWarnings)) sb.AppendLine(healthWarnings);
            sb.AppendLine();

            // === Ch.10 大運逐運（暫停輸出，另有大運命書）===
            // sb.AppendLine("【第十章：大運逐運論斷（百分制評分）】");
            // {
            //     string ageHint = LfAgeTopicHint(currentAge);
            //     if (!string.IsNullOrEmpty(ageHint)) sb.AppendLine(ageHint);
            // }
            string[] branchSSArr = { yBranchSS, mBranchSS, dBranchSS, hBranchSS };
            // foreach (var c in scored)
            // {
            //     string lcSS = LfStemShiShen(c.stem, dStem);
            //     string lcBranchMs = LfBranchHiddenRatio.TryGetValue(c.branch, out var lcBH) && lcBH.Count > 0 ? lcBH[0].stem : "";
            //     string lcBranchSS = !string.IsNullOrEmpty(lcBranchMs) ? LfStemShiShen(lcBranchMs, dStem) : "";
            //     sb.AppendLine($"{c.startAge}-{c.endAge} 歲 大運：{c.stem}{c.branch}（天干{lcSS}·地支{lcBranchSS}）  評分：{c.score} 分（{c.level}）");
            //     sb.AppendLine($"  {LfLuckDesc(c.score, c.level)}");
            //     string events = LfBranchEventsPalace(c.branch, lcBranchSS, branches, branchSSArr, c.startAge);
            //     if (!string.IsNullOrEmpty(events))
            //     {
            //         sb.AppendLine($"  【地支事項】大運地支{c.branch}（{lcBranchSS}）：");
            //         sb.AppendLine($"  {events}");
            //     }
            // }
            // sb.AppendLine();

            // === Ch.11 流年重點（暫停輸出，另有流年命書）===
            // sb.AppendLine("【第十一章：流年重點吉凶】");
            // sb.AppendLine(LfKeyYears(scored, birthYear, yongShenElem, jiShenElem));
            // sb.AppendLine();

            // === Ch.10（原Ch.12）總評 ===
            sb.AppendLine("【第十章：一生命運總評】");

            // 大運行運一覽（含吉凶等級）
            sb.AppendLine("一生大運行運一覽：");
            var curCycleBz = scored.FirstOrDefault(c => currentAge >= c.startAge && currentAge <= c.endAge);
            foreach (var cycle in scored)
            {
                bool isCurrent = currentAge >= cycle.startAge && currentAge <= cycle.endAge;
                string currentMark = isCurrent ? $"  ← 目前行運（{currentAge} 歲）" : "";
                sb.AppendLine($"  {cycle.startAge}-{cycle.endAge} 歲：{cycle.stem}{cycle.branch}（{cycle.level}·{cycle.score}分）{currentMark}");
                if (isCurrent)
                    sb.AppendLine($"    {LfLuckDesc(cycle.score, cycle.level)}");
            }
            sb.AppendLine();

            // 目前行運現況分析
            if (!string.IsNullOrEmpty(curCycleBz.stem))
            {
                sb.AppendLine($"【目前行運現況分析（{curCycleBz.startAge}-{curCycleBz.endAge} 歲 {curCycleBz.stem}{curCycleBz.branch}）】");
                string curStemSS = LfStemShiShen(curCycleBz.stem, dStem);
                string curBrMs = LfBranchHiddenRatio.TryGetValue(curCycleBz.branch, out var curBH) && curBH.Count > 0 ? curBH[0].stem : "";
                string curBrSS = !string.IsNullOrEmpty(curBrMs) ? LfStemShiShen(curBrMs, dStem) : "";
                string curStemElem = KbStemToElement(curCycleBz.stem);
                bool curStemGood = curStemElem == yongShenElem || curStemElem == fuYiElem;
                bool curStemBad = curStemElem == jiShenElem;
                string curStemTrend = curStemGood ? "屬喜用五行，天干助力" : curStemBad ? "屬忌神五行，天干帶阻" : "屬中性五行";
                string curStemEventDesc = curStemSS switch
                {
                    "比肩" => curStemGood ? "自立奮發，同輩互助，合夥共事有利。" : "競爭耗力，同輩牽制，宜各自獨立、防糾紛。",
                    "劫財" => curStemGood ? "積極進取，破舊立新，有偏財機遇。" : "財務競爭激烈，宜防破財耗損、合夥是非。",
                    "食神" => curStemGood ? "才藝展現，口福豐盛，子女緣佳，事業創作機會多。" : "耗洩過度，精力分散，需節制。",
                    "傷官" => curStemGood ? "才華外露，技術精進，適合創業突破舊局。" : "口舌是非多，易與上司對立，宜修身謙遜。",
                    "偏財" => curStemGood ? "偏財運旺，父緣異性緣佳，廣結善緣有助財源。" : "財來財去，易衝動破財，宜謹慎理財。",
                    "正財" => curStemGood ? "財運穩固，努力必有回報，婚姻穩定。" : "財庫受壓，勞而收穫有限，宜節流保守。",
                    "七殺" => curStemGood ? "壓力化為動力，可建功立業，適合競爭激烈的環境。" : "官非壓力大，健康情緒易受損，宜守成防意外。",
                    "正官" => curStemGood ? "名聲地位提升，升遷機會大，婚緣顯現。" : "規範束縛感強，職場壓力重，宜守紀律防小人。",
                    "偏印" => curStemGood ? "偏門學習進修，貴人助力，靈感豐富。" : "思路偏執，食傷受制，宜廣納意見防孤立。",
                    "正印" => curStemGood ? "印綬護身，學業晉升，長輩庇蔭，心靈沉穩。" : "依賴心重，行動力不足，宜主動出擊。",
                    _ => ""
                };
                sb.AppendLine($"  天干 {curCycleBz.stem}（{curStemSS}）{curStemTrend}：{curStemEventDesc}");
                if (!string.IsNullOrEmpty(curBrSS))
                {
                    string curBrElem = LfBranchHiddenRatio.TryGetValue(curCycleBz.branch, out var _) ? KbStemToElement(curBrMs) : "";
                    bool curBrGood = curBrElem == yongShenElem || curBrElem == fuYiElem;
                    bool curBrBad = curBrElem == jiShenElem;
                    string curBrTrend = curBrGood ? "屬喜用五行，地支助力" : curBrBad ? "屬忌神五行，地支帶阻" : "屬中性五行";
                    sb.AppendLine($"  地支 {curCycleBz.branch}（{curBrSS}）{curBrTrend}");
                }
                sb.AppendLine($"  綜合評估：{LfLuckDesc(curCycleBz.score, curCycleBz.level)}");
                string curPalaceEvents = LfBranchEventsPalace(curCycleBz.branch, curBrSS, branches, branchSSArr, curCycleBz.startAge);
                if (!string.IsNullOrEmpty(curPalaceEvents))
                    sb.AppendLine($"  地支六親事項：{curPalaceEvents}");
                sb.AppendLine();
            }

            if (scored.Count > 0)
            {
                var best  = scored.MaxBy(c => c.score)!;
                var worst = scored.MinBy(c => c.score)!;
                sb.AppendLine($"人生最佳期：{best.startAge}-{best.endAge} 歲 {best.stem}{best.branch}（{best.score} 分）");
                sb.AppendLine($"人生考驗期：{worst.startAge}-{worst.endAge} 歲 {worst.stem}{worst.branch}（{worst.score} 分）");
            }
            sb.AppendLine($"財富等級：{LfWealthLevel(caiPct, bodyPct)}  功名等級：{LfFameLevel(pattern)}");
            sb.AppendLine($"命主喜走{yongShenElem}方位（{LfElemDir(yongShenElem)}），從事{LfElemCareer(yongShenElem)}為吉。");
            sb.AppendLine($"人生最重要提醒：善加把握【{yongShenElem}】所代表的人事物。");
            sb.AppendLine($"趨吉避凶：謹慎避免【{jiShenElem}】方向的事情，尤其在中凶/大凶運期間。");
            sb.AppendLine();
            sb.AppendLine(LfBuildWuXingLeiXiang(yongShenElem, jiShenElem));
            sb.AppendLine();

            // === 人生警示事項 ===
            sb.AppendLine("【人生警示事項】");
            sb.AppendLine();
            sb.AppendLine("▍ 小人防範");
            sb.AppendLine(LfXiaoRenAnalysis(yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch, jiShenElem, dmElem));
            sb.AppendLine();
            sb.AppendLine("▍ 官司文書風險");
            sb.AppendLine(LfGuanSiAnalysis(yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch, jiShenElem, dmElem, bodyPct));
            sb.AppendLine();
            sb.AppendLine("▍ 車關時機");
            sb.AppendLine(LfCheGuanAnalysis(yBranch, mBranch, dBranch, hBranch, jiShenElem, dmElem));
            sb.AppendLine();
            sb.AppendLine("▍ 海外發展");
            sb.AppendLine(LfHaiWaiAnalysis(yBranch, mBranch, dBranch, hBranch, yongShenElem, jiShenElem, dmElem, false, default));
            sb.AppendLine();
            sb.AppendLine("▍ 天乙貴人方向");
            {
                var tianYiMapBz = new Dictionary<string, string>
                {
                    {"甲","丑未"},{"戊","丑未"},{"庚","丑未"},
                    {"乙","子申"},{"己","子申"},
                    {"丙","亥酉"},{"丁","亥酉"},
                    {"壬","卯巳"},{"癸","卯巳"},
                    {"辛","午寅"}
                };
                string tianYiBranchesBz = tianYiMapBz.GetValueOrDefault(dStem, "");
                sb.AppendLine($"{dStem} 日主，天乙貴人在：{tianYiBranchesBz}（見此地支方位或行此地支大運，貴人助力最強）");
            }
            sb.AppendLine();
            // 行業建議
            sb.AppendLine("【適合行業綜合建議】");
            {
                var (cfHuangliang, cfYangZhiYin) = KbCalcCareerFlags(
                    yStemSS, yBranchSS, mStemSS, mBranchSS,
                    dBranchSS, hStemSS, hBranchSS, dStem, pattern);
                sb.Append(KbSanmenJobByElem(dmElem, pattern, yongShenElem, cfHuangliang, cfYangZhiYin));
            }
            sb.AppendLine();
            // 居家風水
            sb.AppendLine("【居家風水開運佈置】");
            sb.AppendLine();
            sb.Append(KbSanmenFengShui(
                yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch,
                dmElem, bodyPct, yongShenElem, jiShenElem, wuXing, scored));
            sb.AppendLine();

            sb.AppendLine("-----------------------------------------------------------------");
            sb.AppendLine("命理大師：玉洞子 | 八字命書 v2.3");
            return sb.ToString();
        }

        // === 解析農曆月份 ===
        private static int LfParseLunarMonth(string lunarBirthDate)
        {
            if (string.IsNullOrEmpty(lunarBirthDate)) return 0;

            // 格式一：中文「X年Y月Z日」
            int yearIdx = lunarBirthDate.IndexOf('年');
            int monthIdx = lunarBirthDate.IndexOf('月', yearIdx >= 0 ? yearIdx : 0);
            if (yearIdx >= 0 && monthIdx > yearIdx)
            {
                string monthStr = lunarBirthDate.Substring(yearIdx + 1, monthIdx - yearIdx - 1);
                int chResult = monthStr switch {
                    "一" => 1, "二" => 2, "三" => 3, "四" => 4, "五" => 5, "六" => 6,
                    "七" => 7, "八" => 8, "九" => 9, "十" => 10, "十一" => 11, "十二" => 12,
                    _ => 0
                };
                if (chResult > 0) return chResult;
            }

            // 格式二：「YYYY/MM/DD」或「YYYY-MM-DD」
            var parts = lunarBirthDate.Split('/', '-');
            if (parts.Length >= 3 && int.TryParse(parts[1], out int m2) && m2 >= 1 && m2 <= 12)
                return m2;

            // 格式三：純數字字串（YYYYMMDD 8碼 或 YYYYmDD 7碼）
            string digits = new string(lunarBirthDate.Where(char.IsDigit).ToArray());
            if (digits.Length == 8 && int.TryParse(digits.Substring(4, 2), out int m3) && m3 >= 1 && m3 <= 12)
                return m3;
            if (digits.Length == 7 && int.TryParse(digits.Substring(4, 1), out int m4) && m4 >= 1 && m4 <= 9)
                return m4;

            return 0;
        }

        // === 百分百確信度斷語（條件2-9）===
        private static string LfBaiShengSections(string yBranch, string hBranch, int? birthHour, int? birthMinute,
            int birthYear, int lunarMonth, int gender, string mBranch, string dStem)
        {
            var sb = new StringBuilder();

            // 時辰地支索引 (子=0,...,亥=11)
            var branchOrder = new[] { "子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥" };
            int hIdx = Array.IndexOf(branchOrder, hBranch);
            int yIdx = Array.IndexOf(branchOrder, yBranch);
            if (hIdx < 0) hIdx = 0;
            if (yIdx < 0) yIdx = 0;

            // 計算時段（初/中/末）
            string timeSection = "初";
            if (birthHour.HasValue && birthMinute.HasValue)
            {
                int hourInZhi = birthHour.Value % 2;
                int totalMinutes = hourInZhi * 60 + birthMinute.Value;
                if (totalMinutes < 40) timeSection = "初";
                else if (totalMinutes < 80) timeSection = "中";
                else timeSection = "末";
            }

            // === 條件 2：十二生肖出生時辰論 ===
            var zodiacNames = new Dictionary<string, string>
            {
                {"子","鼠"},{"丑","牛"},{"寅","虎"},{"卯","兔"},{"辰","龍"},{"巳","蛇"},
                {"午","馬"},{"未","羊"},{"申","猴"},{"酉","雞"},{"戌","狗"},{"亥","豬"}
            };
            var zodiacTexts = new Dictionary<string, string[]>
            {
                {"子", new[]{"好如春，到處柳綠花又紅","田園逢，秋熟禾稻積滿倉","泉趣遨遊不知歸","見貓難，退避刑克急須防","春光隨，好處遊覽亦忘歸","鼠落泥窩下，被蛇咬暗自傷","置身倉庫裡，食祿不須求","是非因雀口，禍患自難消","鴉棲庭樹裡，定祿不須求","日落黃昏後，燈前有稻糧","不疾病，多抵因為火來論","五湖波浪闊，天水相接連"}},
                {"丑", new[]{"春風吹大地，草長牧牛肥","生平為溫稱，有志在園林","欲貪岩下草，忽遇虎相遭","芳草萋萋地，閒遊在人間","春生有榮貴，秋令決傷兒","遇蛇刑害至，口舌並官非","遇虎雖無害，凶多吉少成","一犁春雨後，飽臥牧童前","深山獨覓食，獨步見高低","日夕雞棲後，歸來飽臥時","牧童騎昔稱，溪口趁斜輝","覓食東效外，歸時緩緩行"}},
                {"寅", new[]{"出生傷一箭，官訟急須防","扶持如遇貴，金榜決提名","身非晉憑姓，捕虎力難勝","負隅獨立，氣欲吞牛","無端遭暗箭，禍患必臨身","不似山前石，當年李廣奇","嘯來山谷震，動物盡皆逃","見羊難釋放，勢必急須離","禍患隨時至，無憂早須知","園中收日久，猛氣已消除","無所欲一生，惟計在肥豚","耽耽非可比，蔚起接人文"}},
                {"卯", new[]{"人間稱卯日，天上應房星","爰爰承善日，跌跌搗無霜","遇虎春為貴，秋來反不祥","托跡宣王囿，民多有怨詞","婁金猶恐懼，見必主刑沖","壬癸冬生水，逢蛇貴自成","從來人取義，兆必祝禎祥","火多災厄重，不隔主刑傷","托跡逢蒿下，庸庸過一生","最羨霜毛意，尤知玉骨隹","謾誇三窘美，遇上必多刑","和丹成永壽，搗藥不知殃"}},
                {"辰", new[]{"自莊深淵底時變，一舉飛騰上碧天","半天猶矯矯，頃刻兩成雲","若得風雲會，從今變化多","漢室當鴻運，初年大德興","若逢春日桃花浪，千尺波濤湧地高","何時能返轉，一變到青霄","崢嶸頭角滄海裡，不得興雲志未休","茗名劍號雖珠意，井溢泉清本一流","枝頭鴉噪暮，草裡兔游時","昔雨興庭裡，曾勸龍妃宮","蟲名百足毒無比，只怕金雞不怕蛇","騰身滄海裡，正是興雲降雨時"}},
                {"巳", new[]{"他真龍仰月仰頭，表出志氣又顯威","田園遭竹板，正是半凶半吉時","人逢三月春雷動，志欲騰雲上碧霄","忽見狐狸相過害，人離財散浪淘沙","時時三月桃花浪，躍進龍門定可跨","時行方草外，日麗又春和","鴉來先後趕，無禍亦生凶","一生游盡山林興，得看飛鷹正趕跑","一叢芳草萋萋地，滿面祥開樹樹花","一生猶似蛇潛窟，不得風雲志未伸","忽遇蜈蚣咬一口，失了葫蘆沒了皮","突見兩蛇來交會，妻淫子散破盡家"}},
                {"午", new[]{"晚景來到桃樹李下，衣食自無憂","丑宮春風田園內，出入高堂春風間","猛虎傷足下，若不傷頭沒了皮","卯時生得貴人扶，優遊自在有威風","辰宮山岩高，險處興雲吐霧保春風","時逢天狗相侵害，若不傷損自悲傷","午宮入空房，清閒快樂，衣食自然不用求","此限運乖蹇，自歎亦悲悉","申宮一路受坎坷，平生晚得運，早須防","酉宮屋下愛草食，逍遙自在好前途","狗來相侵相，暗中禍患早提防","亥宮一路春色好，四時春景得和氣"}},
                {"未", new[]{"時宮閒遊芳草地，豐衣足食享田園","時宮獨樂田園裡，只恐災害不安心","時宮岩前逢犯虎，大風之燭失光輝","時宮林中傷其箭，此處災難必時來","時宮遇逢春色動，花紅柳綠真光輝","巳宮樹下青青草，不愁衣食不愁饑","山下逢閑火，不見刑沖亦見害","樹碩衣食足，此時財祿非尋常","申宮散步閒遊樂，發達榮華人間稀","時宮安居茅屋內，無憂快樂過時光","忽然虎相逼害，草裡垂珠失色","亥宮江湖風草地，春風桃李正得時"}},
                {"申", new[]{"滿山桃李煞，見樹必攀登","鬱鬱深山裡，誰知飛來一枝箭","飲泉帶結侶，林下自嬉嬉，豐盈又周到","自輕揉土易，果熟自先嘗","采果歸仙洞，尋待後石岩","覓食入林來，忽遇蛇來咬","吊纏終日坐，性優舞猴兒","身穿高嶺霧，猴跳陣溪煙","此身自入流，終日長敢氣富豪","飽食山頭果，思臨澗底泉","與狗相戲耍，生涯即此間","深山塵世隔，終日樂似仙"}},
                {"酉", new[]{"報曉長隨月，知更不畏霜","一生為喜鬥，五德有成名","助妃成婦教，終日任賢勞","若遇狐狸總不吉，無毒亦傷財","在天為帛日，為地應明金","羅浮紅日見，伊屋度初聲","孟蘭門下客，徐徐門度楚函關","當年懷祖逖，起舞獨爭先","敵能收楚項，力可過雅陰","食呼群伴，雄決死驅","一生猶早起，為善在東樂","棄之猶有味，不決麥夫疑"}},
                {"戌", new[]{"吠對唯一犬，百犬自相生","左牽成別號，美獻亦佳名","遇虎身防害，妻兒決損傷","遙河花邊立，長教月底看","河克原本位，象應婁金星","逢蛇傷口舌，暗裡自相爭","吠蘆來祿祿，知是近桃源","雪裡常閑吠，花前只愛眠","有事迎客，無事吠柴門","效勞如馬，功就自當享","胸少容人量，言多妄誕詞","春生離祖命，冬生福祿盈"}},
                {"亥", new[]{"閒遊貴衣食，不求自然來","時園芳草地，出入青山步步春","時生逢虎相損害，縱然聚尾未盡享","山林草荀發，逍遙自在樂其中","遇達春風動，亦有衣祿亦有錢","逢蛇來損害，傷到其足定防妻","被火燒正傷，慎言可行宣晚對","花下偷閒食，天邊送到吉星來","吊捆樹頭上，生死未知正警惶","收養在家中，食祿馬相隨，春風遇時光","被狗相咬害，頭面未免亦有傷","時逢春景處處好，名稱陶豬勝陶豬"}}
            };

            sb.AppendLine("---");
            string zodiac = zodiacNames.TryGetValue(yBranch, out var zn) ? zn : yBranch;
            if (zodiacTexts.TryGetValue(yBranch, out var ztArr) && hIdx < ztArr.Length)
            {
                sb.AppendLine($"特性注意 ：{ztArr[hIdx]}");
            }
            sb.AppendLine();

            // === 條件 3：貴賤論 ===
            // 結構：(命格名, 初命格, 初父母, 中命格, 中父母, 末命格, 末父母)
            var guiJianData = new Dictionary<string, (string name, string initText, string initFather, string midText, string midFather, string finText, string finFather)>
            {
                {"子", ("玉皇星入命",
                    "初年受苦楚，能勤儉持家，末限足財寶", "先克母",
                    "妻賢子孝富貴命，晚年佳，一生近貴人", "父母雙全",
                    "初年榮帶苦，晚年勞碌，六親無依靠，兄弟兩東西", "先克父")},
                {"丑", ("閻王星入命，帶刑克",
                    "性靈心巧，享福無窮，二十年後衣祿積有餘", "父母雙全",
                    "聰明伶俐，初運不如意，晚年享子孫福，三十年發財", "先克父",
                    "特達，常為人排解訟訴，四十歲發財，手藝學術強", "先克母")},
                {"寅", ("白帝星入命",
                    "辛苦冷淡，妻宜硬對，三十八歲發財，末限好", "先克父",
                    "忠直磊落，六親有力，生平有大志，宅舍好風光", "父母雙全",
                    "忠信有始有終，三十六歲發財，治家有法", "先克母")},
                {"卯", ("貴無星入命",
                    "忠良正直，不信鬼神，晚年子女得力", "先克母",
                    "兄弟子女得力，運限佳，身帶微穩暗疾", "父母福壽雙全",
                    "心慈耐勞守本份，初年不遂，中晚年財祿豐盈", "先克父")},
                {"辰", ("人馬元星入命，命帶三刑六害",
                    "四十三歲有一病，對妻宜大，行善可保父母健全", "父母雙全",
                    "待人十分無功，早年坎坷，晚至福興隆", "先克母",
                    "聰明有賢慧，命帶桃花，須遲娶對硬妻，晚榮華", "先克父")},
                {"巳", ("祿元星入命",
                    "四十年前運不遂，晚年逢運食四方財", "先克母",
                    "性急早年憂患，前有見破，亨通晚運行", "父母雙全",
                    "心慈好善，先苦後甜，初年不足，後運必亨通", "先克父")},
                {"午", ("太陽星入命",
                    "想事過人但命不合，常招人急恨，一生運平多災忍", "先克母",
                    "早年得志，志氣超群，有威人敬服，隨處皆傳聞", "父母雙全",
                    "前運不順後運佳，勤儉發達，發達在晚歲早年必遭殃", "先克父")},
                {"未", ("太陰星入命",
                    "生居安然，六親無破損，五福俱全", "父母雙全",
                    "性格寬宏身近貴人，兄弟少力，妻硬子遲，先難後易", "先克母",
                    "配妻要金石方得偕老，中年無多，晚歲家財無", "先克父")},
                {"申", ("海角星入命，貴人不臨，刑克六親",
                    "聰明伶俐百事亨通，兒孫登虎榜，年少奪錦標", "父母雙全",
                    "兄弟不得力，妻宜硬對，六親俱無情，手足不相生", "先克父",
                    "早年破家受苦，三十六歲後運好，晚運重新", "六親冷淡")},
                {"酉", ("忠直好樂",
                    "得有好貴子，百事亨通，夫妻歡度奴婢傍立", "父母雙全",
                    "妻要硬對，前運起倒，後運方吉，能送母終老", "先克父",
                    "兄弟少力，靈敏有主張，好排解爭論", "先克母")},
                {"戌", ("地母星",
                    "手足不得力，三十七歲發財，娶妻宜硬對，子媳遲少", "先克母",
                    "六親無情，娶妻宜晚，早運好中運不足末運好，膽大有智謀", "先克父",
                    "六親無情，妻金石，晚得子，宜手藝術精通，離祖大吉", "父母多災難")},
                {"亥", ("天父星，離祖成家，妻硬對，子遲",
                    "妻子健但命不合，心慈無毒，白手振家聲", "先克母",
                    "乖巧伶俐，卅年無刑克，衣祿好末限好", "父母雙全",
                    "六親刑傷，男克妻女克夫，末運好", "先克父")}
            };

            // guiJianData 貴賤論輸出已移除（準確度不足）

            // === 條件 4：四季王（出生時辰地支對應帝王身體部位）===
            {
                // 依月柱地支判斷季節（寅卯辰=春，巳午未=夏，申酉戌=秋，亥子丑=冬）
                string season4 = "寅卯辰".Contains(mBranch) ? "春"
                    : "巳午未".Contains(mBranch) ? "夏"
                    : "申酉戌".Contains(mBranch) ? "秋"
                    : "冬";

                // 系統一：四季地支對應部位（部位 -> 春夏秋冬地支字串）
                // 同部位多地支用字串包含判斷
                var sys1BodyParts = new[]
                {
                    ("頭",   "子", "午", "申", "巳"),
                    ("肩",   "丑未", "巳未", "子午", "卯酉"),
                    ("腹",   "卯酉", "辰戌", "丑未", "寅申"),
                    ("手",   "巳亥", "卯申", "卯酉", "子午"),
                    ("陰",   "午", "子", "寅", "亥"),
                    ("膝",   "辰戌", "亥酉", "辰戌", "丑未"),
                    ("腳",   "寅申", "寅丑", "巳亥", "辰戌")
                };

                string body4 = "";
                foreach (var (part, chun, xia, qiu, dong) in sys1BodyParts)
                {
                    string seasonBranches = season4 switch { "春" => chun, "夏" => xia, "秋" => qiu, _ => dong };
                    if (seasonBranches.Contains(hBranch)) { body4 = part; break; }
                }

                // 各部位斷語（季節專屬）
                var bodyTexts4 = new Dictionary<string, Dictionary<string, string>>
                {
                    {"頭", new Dictionary<string, string>{
                        {"春","一世永無憂，求官皆有應，衣祿自然足，男近王侯，女嫁好兒郎"},
                        {"夏","白手振英豪，財祿足，出入近貴人，女只恐夫妻重"},
                        {"秋","財祿主高遷，自身憂不足，男享福應晚年，女多嗟歎有口難言"},
                        {"冬","富貴福壽高，子貴妻賢，男貴至公侯，女命誥封，只恐浮生百事憂"}
                    }},
                    {"肩", new Dictionary<string, string>{
                        {"春","衣祿足，奴婢兩過立，田園廣，初年命苦，晚主福綿綿"},
                        {"夏","近貴人，聰明多巧智，名聲四海揚，男多才干，女多刑克夫子不親"},
                        {"秋","縱然早顛倒，末主衣祿，男財豐富貴足，女享福壽添"},
                        {"冬","艱苦備嘗，衣食難繼，晚景福有餘，早年勞苦尾景祿齊全"}
                    }},
                    {"腹", new Dictionary<string, string>{
                        {"春","衣祿自然足，文武兩邊立，男富貴足，女清閒祖業興福祿"},
                        {"夏","財帛自然足，田園富，男福壽，女作夫人，兒孫得力"},
                        {"秋","夫妻難結果，到老無兒孫，男離祖晚享福，女有刑克六親稀"},
                        {"冬","財祿不須疑，想財即至，男離祖借姓防克六親，女良人食祿足"}
                    }},
                    {"手", new Dictionary<string, string>{
                        {"春","衣祿不長久，女多命苦，男防災厄，急急修行好"},
                        {"夏","富貴有黃金，財帛自然進，男只恐損六親，女多刑克憂愁勞心"},
                        {"秋","福田早好修，兒女無憂，男福祿昌，女多清福產兒孫強"},
                        {"冬","金貴有錢財，聰明智慧，男有名聲財祿不勞，女主福壽子孫英豪"}
                    }},
                    {"陰", new Dictionary<string, string>{
                        {"春","富貴足珍珠，中年衣祿，永聚黃金，男乾坤志貴人身，女兒孫祿滿堂"},
                        {"夏","春花待雨開，嫩柳狂風怕，勸君早修行，鴛鴦不獨宿，飛燕自成雙"},
                        {"秋","室內積千金，子媳宜遲，男增福壽晚景昌，女主聰明衣祿福綿綿"},
                        {"冬","平素大勞心，晚年財發達，早歲艱辛"}
                    }},
                    {"膝", new Dictionary<string, string>{
                        {"春","做事無利益，初年平，中年衣祿，男勞碌心不足，一生亦平安修得是福"},
                        {"夏","白手振家庭，豐衣財祿足，男出外吉經營通，女出家方成福"},
                        {"秋","作事有差失，走盡天涯路，妻無子不應，男克親離祖，女克夫守孤身"},
                        {"冬","命硬過如鐵，婚姻配金石方得偕老"}
                    }},
                    {"腳", new Dictionary<string, string>{
                        {"春","修行卻是福，不宜居祖屋，男移過祖娶妻再續，女嫁二夫有子享福"},
                        {"夏","少年多勞碌，奔走江湖，計謀皆反復，四十不惑後停交中年運方可得"},
                        {"秋","一生長享福，六親天然份，外出享榮華，男離祖多發福，女只恐守孤身"},
                        {"冬","一生主勞碌，衣食朝朝有，離祖方成福，男克六親白手成家，女夫子不親"}
                    }}
                };

                if (!string.IsNullOrEmpty(body4) && bodyTexts4.TryGetValue(body4, out var bDict4) && bDict4.TryGetValue(season4, out var bText4))
                {
                    sb.AppendLine(bText4);
                    sb.AppendLine();
                }
            }

            // === 條件 5：小兒關煞（出生季節 × 時辰）===
            {
                // 依月柱地支判斷季節
                string season5 = "寅卯辰".Contains(mBranch) ? "春"
                    : "巳午未".Contains(mBranch) ? "夏"
                    : "申酉戌".Contains(mBranch) ? "秋"
                    : "冬";

                // 結構：(季節, 時辰地支, 關煞名稱)，同時辰可有多條
                var keShaTable = new List<(string season, string branches, string name)>
                {
                    // 春季
                    ("春","丑","閻王關"),("春","辰","上盆關"),("春","子亥","急腳關"),
                    ("春","辰","將軍箭"),("春","巳","落井關"),("春","寅申","夜啼關"),
                    ("春","子巳","浴盆關"),("春","子","三丘煞"),("春","酉","生命關"),
                    // 夏季
                    ("夏","辰","閻王關"),("夏","申","四季關"),("夏","未","上盆關"),
                    ("夏","子","生命關"),("夏","卯","急腳關"),("夏","子","將軍箭"),
                    ("夏","酉","夜啼關"),("夏","未","深水關"),
                    // 秋季
                    ("秋","子","閻王關"),("秋","亥","四季關"),("秋","戌","上盆煞"),
                    ("秋","寅","急腳關"),("秋","酉","落井關"),("秋","未","三丘煞"),
                    ("秋","卯","生命關"),("秋","戌","沐浴關"),
                    // 冬季
                    ("冬","卯","閻王關"),("冬","寅戌","四季關"),("冬","子","上盆煞"),
                    ("冬","丑","四季關"),("冬","辰戌","急腳關"),("冬","申巳亥","將軍箭"),
                    ("冬","卯","夜啼關")
                };

                var found5 = keShaTable
                    .Where(r => r.season == season5 && r.branches.Contains(hBranch))
                    .Select(r => r.name)
                    .ToList();

                if (found5.Count > 0)
                {
                    string shaNames = string.Join("、", found5);
                    sb.AppendLine($"帶有：{shaNames}。");
                    if (found5.Count >= 2)
                        sb.AppendLine("多重關煞疊加，幼年凶險較重，宜提早化解。");
                    else
                        sb.AppendLine("幼年需多加注意身體平安。");
                }
                else
                {
                    sb.AppendLine("幼年平順。");
                }
                sb.AppendLine();
            }

            // === 條件 6：子女送終 ===
            var palaceNames = new[] { "閉","開","收","成","危","破","執","定","平","滿","除","建" };
            // 閉宮對應時支起始索引 = (yIdx + 2) % 12
            int biGongStart = (yIdx + 2) % 12;
            int palaceIdx = (hIdx - biGongStart + 12) % 12;
            string palace = palaceNames[palaceIdx];
            var palaceTexts = new Dictionary<string, string>
            {
                {"建", "雖有子女，歸仙未必全，惟留一二，送柩到墳前"},
                {"除", "子孫送君壽，一生永無憂，衣食自有餘，家業得成就，父母壽千秋"},
                {"滿", "夫妻永百年，子孫送君壽，萬頃好良田，財穀豐盈積，人稱福壽全"},
                {"平", "子孫送君壽，一世永無虧，若再加陰德，富貴不須求"},
                {"定", "富貴雙全命，子孫送君壽，為官多善政，子孫名聲業"},
                {"執", "衣祿自然足，子孫送君壽，一生多享福，家業自操持"},
                {"破", "煩惱心頭掛，有子送君壽，隨份隨緣過，孫賢與子尚"},
                {"危", "家業不大富，出外多近貴，末運始興隆，子孫送君壽"},
                {"成", "富貴自分明，子孫送君壽，子振門庭，待至曾元輩，衣冠自業榮"},
                {"收", "衣祿不須求，子孫送終去，雖多出外遊，子孫有清福"},
                {"開", "不怕沒銀錢，縱有子孫在，送壽失身邊，食祿隨君份"},
                {"閉", "衣祿自然豐，子孫送君壽，福祿永無窮，至老相伴隨"}
            };

            string palaceText = palaceTexts.TryGetValue(palace, out var pt) ? pt : "";
            sb.AppendLine(palaceText);
            sb.AppendLine();

            // === 條件 7：查流年 ===
            int currentYear = DateTime.Today.Year;
            int vage = currentYear - birthYear + 1;
            if (vage > 0)
            {
                int pos1 = (vage - 1) % 12;
                var sys1Stars = new[] {"太歲","太陽","喪門","太陰","官符","死符","歲破","龍德","白虎","福德","吊客","病符"};
                var sys1Texts = new Dictionary<string, string>
                {
                    {"太歲", "太歲當頭照，諸神不敢當，自己有刑克，須防克親人"},
                    {"太陽", "太陽星吉照，求財得珠寶，陰陽須和合，福祿自然來"},
                    {"喪門", "喪門入命來，內外孝服來，破財並刑克，災難又重臨"},
                    {"太陰", "太陰入命來，見喜又見財，貴人相接引，萬事必和諧"},
                    {"官符", "官符入命來，禍患必無常，破財並外孝，祈保免災殃"},
                    {"死符", "死符入命來，病疾主有災，若無官符事，膿血積生害"},
                    {"歲破", "歲破入命來，孝服哭哀哀，父母有刑克，損妻又損財"},
                    {"龍德", "龍德貴人生，發財旺人丁，小災不為害，家宅最光明"},
                    {"白虎", "白虎星辰賽，春夏防災厄，秋冬平平過，祈保無邪惑"},
                    {"福德", "福星入命遊，一歲進田牛，添丁又進契，凡事吉安康"},
                    {"吊客", "吊客入命凶，內外孝服重，破財並損口，祈保免災殃"},
                    {"病符", "病符入命來，疾病必沾身，時運多顛倒，要好必求神"}
                };
                var sys1LuckBad = new Dictionary<string, string>
                {
                    {"太歲","凶"},{"太陽","吉"},{"喪門","凶"},{"太陰","吉"},{"官符","凶"},{"死符","凶"},
                    {"歲破","凶"},{"龍德","吉"},{"白虎","凶"},{"福德","吉"},{"吊客","凶"},{"病符","凶"}
                };

                int pos2 = (vage - 1) % 12;
                var sys2Stars = new[] {"天哭","福德","黑煞","文昌","火孛","暗曜","天掃","太陰","天喜","無常","龍德","錦繡"};
                var sys2Texts = new Dictionary<string, string>
                {
                    {"天哭", "主孝服官非口舌、疾病、災厄，有喜免哭，無喜要防"},
                    {"福德", "主吉多凶少，官非口舌盡消除，有喜無喜俱美"},
                    {"黑煞", "主官非口舌，刑克，喪服，有喜可押身，無喜當防"},
                    {"文昌", "主有吉兆，士者名標金榜，庶人招財，女人有喜"},
                    {"火孛", "主官非口舌，破財，刑克孝服，損六畜，小人侵害，有喜可防"},
                    {"暗曜", "主刑克破財損六畜，防盜賊，有喜可免，無喜當保"},
                    {"天掃", "主有孝服，官非口舌疾病，有不測之禍，有喜可免，無喜當防"},
                    {"太陰", "主有喜事臨門，家中百事俱吉，女人有喜定是男兒，無喜亦吉"},
                    {"天喜", "主有喜慶，家道吉祥四季得財，但主有血刃之災，女人有喜是男兒"},
                    {"無常", "主有不測之災及刑克孝服，有喜可免，無喜須防"},
                    {"龍德", "主喜慶臨門家中吉慶百事享通，縱有血刃之災亦無防"},
                    {"錦繡", "主有喜，事事遂意，女人有喜定是男兒"}
                };
                var sys2LuckBad = new Dictionary<string, string>
                {
                    {"天哭","凶"},{"福德","吉"},{"黑煞","凶"},{"文昌","吉"},{"火孛","凶"},{"暗曜","凶"},
                    {"天掃","凶"},{"太陰","吉"},{"天喜","吉帶血刃"},{"無常","凶"},{"龍德","吉"},{"錦繡","吉"}
                };

                string star1 = sys1Stars[pos1];
                string star2 = sys2Stars[pos2];
                string lb1 = sys1LuckBad.TryGetValue(star1, out var l1) ? l1 : "";
                string lb2 = sys2LuckBad.TryGetValue(star2, out var l2) ? l2 : "";
                string t1 = sys1Texts.TryGetValue(star1, out var tx1) ? tx1 : "";
                string t2 = sys2Texts.TryGetValue(star2, out var tx2) ? tx2 : "";

                // 第七條（查流年）僅用於流年命書，此處不輸出
            }

            // === 條件 8：讀書格 ===
            // 月支直接對應農曆月份：寅=1,卯=2,辰=3,巳=4,午=5,未=6,申=7,酉=8,戌=9,亥=10,子=11,丑=12
            var mBranchToMonth = new Dictionary<string, int>
            {
                {"寅",1},{"卯",2},{"辰",3},{"巳",4},{"午",5},{"未",6},
                {"申",7},{"酉",8},{"戌",9},{"亥",10},{"子",11},{"丑",12}
            };
            int lunarMonthFromBranch = mBranchToMonth.TryGetValue(mBranch, out var lmb) ? lmb : 0;
            if (lunarMonthFromBranch >= 1)
            {
                // 三合組
                string group8;
                int groupCol;
                if ("寅午戌".Contains(yBranch)) { group8 = "寅午戌"; groupCol = 0; }
                else if ("申子辰".Contains(yBranch)) { group8 = "申子辰"; groupCol = 1; }
                else if ("巳酉丑".Contains(yBranch)) { group8 = "巳酉丑"; groupCol = 2; }
                else { group8 = "亥卯未"; groupCol = 3; }

                // 月份對應格局 (1-based月, 4 columns)
                var monthTable = new string[,]
                {
                    // 寅午戌, 申子辰, 巳酉丑, 亥卯未
                    {"建","破","向","背"},   // 1
                    {"背","向","合","空"},   // 2
                    {"背","向","合","空"},   // 3
                    {"背","向","建","破"},   // 4
                    {"空","合","背","向"},   // 5
                    {"空","合","背","向"},   // 6
                    {"破","建","背","向"},   // 7
                    {"向","背","空","合"},   // 8
                    {"向","背","空","合"},   // 9
                    {"向","背","破","建"},   // 10
                    {"合","空","向","背"},   // 11
                    {"合","空","向","背"}    // 12
                };
                string geju8 = monthTable[lunarMonthFromBranch - 1, groupCol];
                var geju8Texts = new Dictionary<string, string>
                {
                    {"建", "有學堂。生居建學貌堂堂，習讀文章事事昌，不是為官做宰相，也做巧性俊兒郎"},
                    {"向", "有學堂。命中向學近書堂，必主文章藝術強，不作秀才棟樑命，也作儒家出類郎"},
                    {"合", "有學堂。命中向學近書堂，文章藝術強"},
                    {"背", "無學堂。生來空破皆文章，性情識字不成行，老來思量年少事，讀書不曉怨爹娘"},
                    {"空", "無學堂。同背"},
                    {"破", "無學堂。同背"}
                };
                string geju8Text = geju8Texts.TryGetValue(geju8, out var gt) ? gt : "";

                sb.AppendLine($"{geju8}格——{geju8Text}");
                sb.AppendLine();
            }

            // === 條件 9：論時看孤雙 ===
            string group9Name, group9Common, group9TimeSect;
            if ("子午卯酉".Contains(hBranch))
            {
                group9Name = "四時高（子午卯酉）";
                group9Common = "為人清秀主英豪，多者兄弟難為伴，父母緣份薄。";
                group9TimeSect = (timeSection == "中")
                    ? "富貴招，財祿榮華須自在，名望通天得勢高。"
                    : "遠末無依靠，孤立自立。";
            }
            else if ("寅申巳亥".Contains(hBranch))
            {
                group9Name = "四時強（寅申巳亥）";
                group9Common = "為人聰秀有文章，自然享福靠自己。";
                group9TimeSect = (timeSection == "中")
                    ? "兄弟較多，父母親疏無依靠。"
                    : "兄弟不多，父母親疏無依靠。";
            }
            else
            {
                // 辰戌丑未
                group9Name = "四時孤（辰戌丑未）";
                group9Common = "兄弟無依靠，祖業不守，受奔波命。";
                string group9Parent = (timeSection == "中") ? "先克父。" : "先亡母。";
                string group9OutHome = (gender == 1) ? "男命有出家或學道傾向。" : "女命有出家為尼傾向。";
                group9TimeSect = group9Parent + group9OutHome;
            }

            sb.AppendLine(group9Common);
            sb.AppendLine(group9TimeSect);
            sb.AppendLine();

            return sb.ToString().TrimEnd('\n', '\r') + "\n";
        }

        // === ShiWenSection (審時聞切) ===
        // === 按季節過濾 MonthInfluence / MaleChart 文字 ===
        // 格式：「生於春季（...），文字；生於夏季（...），文字。」用「；」分段，只保留當前季節
        private static string LfFilterSeasonText(string? text, string seasonChar)
        {
            if (string.IsNullOrWhiteSpace(text)) return text ?? "";
            if (!text.Contains("生於") || !text.Contains("季")) return text;
            var parts = text.Split('；');
            var kept = new List<string>();
            foreach (var part in parts)
            {
                var m = System.Text.RegularExpressions.Regex.Match(part.TrimStart(), @"^生於([春夏秋冬])季");
                if (m.Success)
                {
                    if (m.Groups[1].Value == seasonChar) kept.Add(part.Trim());
                }
                else
                {
                    kept.Add(part.Trim());
                }
            }
            if (kept.Count == 0) return text;
            return string.Join("；", kept);
        }

        private static string LfShiWenSection(string timeBranch, int? birthHour, int? birthMinuteRaw, int gender)
        {
            if (!birthHour.HasValue || !birthMinuteRaw.HasValue)
                return "（時辰精確刻分未提供，審時聞切略去，建議當面確認時辰後補驗）\n";

            int hourInZhi = birthHour.Value % 2;
            int totalMinutes = hourInZhi * 60 + birthMinuteRaw.Value;
            int totalQuarter = (totalMinutes / 15) + 1;
            bool isUpperFour = totalQuarter <= 4;
            int relativeQuarter = isUpperFour ? totalQuarter : totalQuarter - 4;
            int personCount = relativeQuarter;
            string timeSection = isUpperFour ? "時初" : "時末";

            bool hasMark = (gender == 1 && isUpperFour) || (gender == 2 && !isUpperFour);
            string markLocation = relativeQuarter switch { 1 => "臉上", 2 => "身上", 3 => "手上", 4 => "腳上", _ => "身上" };
            bool isYangBranch = "子寅辰午申戌".Contains(timeBranch);
            string lookLike = gender == 1 ? (isYangBranch ? "像母親" : "像父親") : (isYangBranch ? "像父親" : "像母親");
            string personality = "。";
            string markText = hasMark ? $"依古法推算，您在「{markLocation}」應有胎記或疤痕。" : "依四時定數，您天生外觀應無明顯胎記。";

            string mingshuDetail = "";
            if ("辰戌丑未".Contains(timeBranch))
            {
                mingshuDetail = "辰戌丑未四時孤，不妨父母少親疏。";
                mingshuDetail += isUpperFour ? "時初時末先亡母。" : "時正者先亡父。";
                mingshuDetail += "兄弟無依靠，祖業難守，男宜僧道女宜姑。";
            }

            var sb = new StringBuilder();
            sb.AppendLine("【審時聞切 · 四時定數】");
            sb.AppendLine($"您生於 {timeBranch}時之{timeSection}（第{totalQuarter}刻）。外貌個性{lookLike}{personality}");
            sb.AppendLine($"印記印證：{markText}");
            if (!string.IsNullOrEmpty(mingshuDetail))
            {
                sb.AppendLine();
                sb.AppendLine("【時柱驗明】");
                sb.AppendLine(mingshuDetail);
            }
            return sb.ToString().TrimEnd('\n', '\r') + "\n";
        }

        // === 十神縮寫 ===
        private static string LfShiShenAbbr(string stem, string dStem)
        {
            return LfStemShiShen(stem, dStem) switch
            {
                "比肩" => "比", "劫財" => "劫", "食神" => "食", "傷官" => "傷",
                "偏財" => "才", "正財" => "財", "七殺" => "殺", "正官" => "官",
                "偏印" => "梟", "正印" => "印", _ => ""
            };
        }

        // 五行 → 十神群組簡稱（印/比/食/財/官）
        private static string LfElemSsGroup(string elem, string dmElem)
        {
            if (elem == dmElem) return "比";
            if (elem == LfGenByElem.GetValueOrDefault(dmElem, "")) return "印";
            if (elem == LfElemGen.GetValueOrDefault(dmElem, "")) return "食";
            if (elem == LfElemOvercome.GetValueOrDefault(dmElem, "")) return "財";
            if (elem == LfElemOvercomeBy.GetValueOrDefault(dmElem, "")) return "官";
            return "";
        }

        private static string LfFmtHidden(string branch, string dStem)
        {
            if (!LfBranchHiddenRatio.TryGetValue(branch, out var stems)) return "";
            return string.Join("", stems.Select(s => LfShiShenAbbr(s.stem, dStem) + s.stem));
        }

        // === 旺相休囚死 ===
        private static (string wang, string xiang, string xiu, string qiu, string si) LfGetWangXiang(string mBranch)
        {
            string monthElem = LfBranchHiddenRatio.TryGetValue(mBranch, out var bh) && bh.Count > 0
                ? KbStemToElement(bh[0].stem) : "土";
            return (
                monthElem,
                LfElemGen.GetValueOrDefault(monthElem, ""),
                LfGenByElem.GetValueOrDefault(monthElem, ""),
                LfElemOvercome.GetValueOrDefault(monthElem, ""),
                LfElemOvercomeBy.GetValueOrDefault(monthElem, "")
            );
        }

        // === 玉洞子命書 v2.0（十六章版）===
        private static string LfBuildYudongziReportV2(
            string yStem, string yBranch, string mStem, string mBranch,
            string dStem, string dBranch, string hStem, string hBranch,
            string yStemSS, string mStemSS, string hStemSS,
            string yBranchSS, string mBranchSS, string dBranchSS, string hBranchSS,
            string yNaYin, string mNaYin, string dNaYin, string hNaYin,
            string dmElem, Dictionary<string, double> wuXing, double bodyPct, string bodyLabel,
            string season, string seaLabel, string pattern, string yongShenElem, string fuYiElem,
            string yongReason, string jiShenElem,
            List<(string stem, string branch, string liuShen, int startAge, int endAge, int score, string level)> scored,
            int gender, int birthYear, int? birthMonth, int? birthDay, int? birthHour, int? birthMinute, int lunarMonth,
            bool hasZiwei, JsonElement palacesYdz, string mingGongStars, string mingZhu, string shenZhu, string wuXingJuText,
            string ziweiMing, string starDescMing, string ziweiFullContent, HashSet<string> chartStars,
            string ziweiOff, string offStars, string ziweiWlt, string wltStars,
            string ziweiSps, string spsStars, string ziweiHlt, string hltStars,
            Dictionary<string, (string pal, string txt)> siHua,
            string nianSiHuaXing, string nianGan,
            string siHuaLuPalace, string siHuaLu,
            string siHuaQuanPalace, string siHuaQuan,
            string siHuaKePalace, string siHuaKe,
            string siHuaJiPalace, string siHuaJi,
            BaziDayPillarReading? kb,
            List<BaziDirectRule>? zhongyuanRules = null,
            string ziweiGeJuContent = "",
            Dictionary<string, string>? doubleDescs = null,
            Dictionary<string, string>? minorDescs = null,
            Dictionary<string, string>? allPalaceStarDescs = null,
            string starDescOff = "",
            string starDescWlt = "",
            string starDescSps = "",
            string starDescHlt = "",
            string ziweiParStar = "",
            string ziweiPar = "",
            string ziweiCldStar = "",
            string ziweiCld = "",
            string userName = "",
            CalendarDbContext? calDb = null)
        {
            var sb = new StringBuilder();
            string genderText = gender == 1 ? "男（乾造）" : "女（坤造）";
            var branches = new[] { yBranch, mBranch, dBranch, hBranch };
            int currentAge = DateTime.Today.Year - birthYear;
            // 個人化：命主稱謂（先生/小姐）
            string genderSuffix = gender == 1 ? "先生" : "小姐";
            string personRef = string.IsNullOrEmpty(userName) ? "命主" : (userName + genderSuffix);
            // 當前大運
            var curCycle = scored.FirstOrDefault(c => c.startAge <= currentAge && c.endAge >= currentAge);
            string curCycleNote = curCycle != default
                ? $"  目前行運：{curCycle.stem}{curCycle.branch} 大運（{curCycle.startAge}～{curCycle.endAge} 歲）"
                : "";

            // 展開十神短稱為全稱，供 ZrApplies 使用
            yStemSS   = KbExpandLiuShen(yStemSS);
            mStemSS   = KbExpandLiuShen(mStemSS);
            hStemSS   = KbExpandLiuShen(hStemSS);
            yBranchSS = KbExpandLiuShen(yBranchSS);
            mBranchSS = KbExpandLiuShen(mBranchSS);
            dBranchSS = KbExpandLiuShen(dBranchSS);
            hBranchSS = KbExpandLiuShen(hBranchSS);

            // 中原盲派直斷輔助 helpers
            var zRules = zhongyuanRules ?? new List<BaziDirectRule>();
            var allStems4 = new[] { yStem, mStem, dStem, hStem };
            var allBranches4 = new[] { yBranch, mBranch, dBranch, hBranch };
            var repeatedStems    = allStems4.GroupBy(s => s).Where(g => g.Count() >= 3).Select(g => g.Key).ToHashSet();
            var repeatedBranches = allBranches4.GroupBy(b => b).Where(g => g.Count() >= 3).Select(g => g.Key).ToHashSet();
            void AppZrList(StringBuilder s, IEnumerable<BaziDirectRule> list, string header)
            {
                var arr = list.ToList();
                if (arr.Count == 0) return;
                s.AppendLine(header);
                foreach (var r in arr)
                    s.AppendLine($"• {r.Content}");
                s.AppendLine();
            }

            // 方案C：條件明確者自動判斷，其餘全部輸出
            string yPillarZ = yStem + yBranch;
            string mPillarZ = mStem + mBranch;
            string dPillarZ = dStem + dBranch;
            string hPillarZ = hStem + hBranch;
            var brZ = new[] { yBranch, mBranch, dBranch, hBranch };
            bool BrHas(string b) => brZ.Contains(b);
            var stemPillarNoSibling = new Dictionary<string, string[]>
            {
                ["甲"] = new[] { "甲申", "庚申" }, ["乙"] = new[] { "乙酉", "辛酉" },
                ["丙"] = new[] { "丙子", "壬子" }, ["丁"] = new[] { "丁亥", "癸亥" },
                ["戊"] = new[] { "戊寅", "甲寅" }, ["己"] = new[] { "己卯", "乙卯" },
                ["庚"] = new[] { "庚寅", "丙寅" },
                ["壬"] = new[] { "癸未", "丙午" }, ["癸"] = new[] { "丙午", "癸未" },
            };
            var tenStems = new[] { "甲","乙","丙","丁","戊","己","庚","辛","壬","癸" };
            var clashPairs = new HashSet<(string, string)>
            {
                ("子","午"),("午","子"),("卯","酉"),("酉","卯"),
                ("寅","申"),("申","寅"),("巳","亥"),("亥","巳"),
                ("辰","戌"),("戌","辰"),("丑","未"),("未","丑")
            };
            bool ZrApplies(BaziDirectRule r)
            {
                string cond = r.Condition;
                switch (r.RuleType)
                {
                    case "SiblingInfo":
                    {
                        // 日干年月柱系列：先判斷條件屬於哪個日干，不對應則直接排除
                        string? ruleStem = tenStems.FirstOrDefault(s => cond.StartsWith(s + "日年月柱"));
                        if (ruleStem != null)
                        {
                            if (ruleStem != dStem) return false;
                            return stemPillarNoSibling.TryGetValue(dStem, out var tgP) &&
                                   (tgP.Contains(yPillarZ) || tgP.Contains(mPillarZ));
                        }
                    }
                        if (cond == "丁丑丁未日時無兄弟")
                            return dPillarZ == "丁丑" || dPillarZ == "丁未";
                        if (cond == "戊寅己卯日克兄弟")
                            return dPillarZ == "戊寅" || dPillarZ == "己卯";
                        if (cond == "日月兩柱干同支沖無兄弟")
                            return mStem == dStem && clashPairs.Contains((mBranch, dBranch));
                        if (cond == "年干為殺非長子")
                            return yStemSS == "七殺";
                        if (cond == "月干為殺非長子")
                            return mStemSS == "七殺";
                        if (cond == "月支為殺定是長子")
                            return mBranchSS == "七殺" || mBranchSS == "正官";
                        if (cond == "年干比劫非長子" || cond == "年干比劫被合兄弟送養")
                            return yStemSS == "比肩" || yStemSS == "劫財";
                        if (cond == "正官在月干非長子")
                            return mStemSS == "正官";
                        if (cond == "年干正官為長子")
                            return yStemSS == "正官";
                        if (cond == "正官正印正財均透干定是長子")
                        {
                            var ss4 = new[] { yStemSS, mStemSS, hStemSS };
                            return ss4.Contains("正官") && ss4.Contains("正印") && ss4.Contains("正財");
                        }
                        if (cond == "月柱甲辰乙未截腳兄弟有損")
                            return mPillarZ == "甲辰" || mPillarZ == "乙未";
                        if (cond == "身弱干無比劫地支比劫被沖無兄弟")
                        {
                            var sStemBJ = new[] { yStemSS, mStemSS, hStemSS };
                            bool stemHasBJ = sStemBJ.Any(s => s == "比肩" || s == "劫財");
                            return bodyPct < 50 && !stemHasBJ;
                        }
                        if (cond == "年干支皆偏財幼年多為養子")
                            return yStemSS == "偏財" && yBranchSS == "偏財";
                        if (cond == "男命七殺旺逢比劫有兄無弟")
                        {
                            if (gender != 1) return false;
                            var aSS28 = new[] { yStemSS, mStemSS, hStemSS, yBranchSS, mBranchSS, dBranchSS, hBranchSS };
                            return aSS28.Any(s => s == "七殺") && aSS28.Any(s => s == "比肩" || s == "劫財");
                        }
                        if (cond == "柱中比肩偏財旺獨生子")
                        {
                            var aSS29 = new[] { yStemSS, mStemSS, hStemSS, yBranchSS, mBranchSS, dBranchSS, hBranchSS };
                            return aSS29.Count(s => s == "比肩") >= 2 || aSS29.Count(s => s == "偏財") >= 2;
                        }
                        if (cond == "偏官偏印偏財重疊定是庶子")
                        {
                            var aSS30 = new[] { yStemSS, mStemSS, hStemSS, yBranchSS, mBranchSS, dBranchSS, hBranchSS };
                            return aSS30.Any(s => s == "七殺") && aSS30.Any(s => s == "偏印") && aSS30.Any(s => s == "偏財");
                        }
                        return true;

                    case "CareerInfo":
                    {
                        var stAll = new[] { yStem, mStem, hStem }; // 年月時干（不含日干）
                        var asSS7 = new[] { yStemSS, mStemSS, hStemSS, yBranchSS, mBranchSS, dBranchSS, hBranchSS };
                        // 身強弱職業分類（三選一）
                        if (cond == "身強財弱宜工業")   return bodyPct >= 60;
                        if (cond == "身財兩停宜商業")   return bodyPct >= 45 && bodyPct < 60;
                        if (cond == "身弱財多宜服務業") return bodyPct < 45;
                        // 辰戌公吏
                        if (cond == "四柱地支辰戌公吏獄官")
                            return BrHas("辰") || BrHas("戌");
                        // 天醫宜醫學（四長生地支：辰巳戌亥）
                        if (cond == "柱中辰巳戌亥天醫星旺宜醫學")
                            return BrHas("辰") || BrHas("巳") || BrHas("戌") || BrHas("亥");
                        // 身弱印旺侍奉人
                        if (cond == "身弱印旺侍奉人之工作")
                        {
                            string yinElem = dmElem switch { "木"=>"水","火"=>"木","土"=>"火","金"=>"土","水"=>"金",_=>"" };
                            return bodyPct < 45 && wuXing.GetValueOrDefault(yinElem, 0) >= 25;
                        }
                        // 年支巳日時申
                        if (cond == "年支巳日時支申長年奔波行業")
                            return yBranch == "巳" && (dBranch == "申" || hBranch == "申");
                        // 天干庚甲地支寅申
                        if (cond == "天干庚甲地支寅申商販郵遞")
                            return (stAll.Contains("庚") || stAll.Contains("甲")) && (BrHas("寅") || BrHas("申"));
                        // 甲丙戊年干+丁+戌亥重 僧道
                        if (cond == "年干甲丙戊丁火地支重戌亥僧道")
                        {
                            bool yOk = yStem == "甲" || yStem == "丙" || yStem == "戊";
                            bool hasD = new[] { mStem, dStem, hStem }.Contains("丁");
                            int xuHai = brZ.Count(b => b == "戌" || b == "亥");
                            return yOk && hasD && xuHai >= 2;
                        }
                        // 華蓋（神殺計算複雜，暫保留顯示）
                        if (cond == "華蓋逢空多為五術之人" || cond == "華蓋太極臨戌亥多為五玄之人")
                            return true;
                        // 女傷官支旺藏財
                        if (cond == "女子傷官在支旺藏財風塵")
                        {
                            if (gender != 2) return false;
                            var brSS4 = new[] { yBranchSS, mBranchSS, dBranchSS, hBranchSS };
                            return brSS4.Any(s => s == "傷官") && brSS4.Any(s => s == "正財" || s == "偏財");
                        }
                        // 水土農民
                        if (cond == "四柱水土重重為農民")
                            return (wuXing.GetValueOrDefault("水", 0) + wuXing.GetValueOrDefault("土", 0)) >= 60;
                        // 正星/偏星
                        if (cond == "正星在柱中多旺從公家政府")
                        {
                            int zheng = asSS7.Count(s => s=="正財"||s=="正官"||s=="正印"||s=="食神"||s=="比肩");
                            int pian  = asSS7.Count(s => s=="偏財"||s=="七殺"||s=="偏印"||s=="傷官"||s=="劫財");
                            return zheng > pian;
                        }
                        if (cond == "偏星在柱中多旺從偏業")
                        {
                            int zheng = asSS7.Count(s => s=="正財"||s=="正官"||s=="正印"||s=="食神"||s=="比肩");
                            int pian  = asSS7.Count(s => s=="偏財"||s=="七殺"||s=="偏印"||s=="傷官"||s=="劫財");
                            return pian > zheng;
                        }
                        // 辛丁巳支酉亥未 酒店
                        if (cond == "柱中辛丁巳支酉亥未宜酒店生意")
                            return (stAll.Contains("辛") || stAll.Contains("丁")) &&
                                   (BrHas("巳") || BrHas("酉") || BrHas("亥") || BrHas("未"));
                        // 沖類（職業/居地變遷）
                        if (cond == "子午卯酉沖地域變遷職業不變")
                            return (BrHas("子") && BrHas("午")) || (BrHas("卯") && BrHas("酉"));
                        if (cond == "寅申巳亥沖居位地和職業都改變")
                            return (BrHas("寅") && BrHas("申")) || (BrHas("巳") && BrHas("亥"));
                        if (cond == "辰戌丑未沖職業改變居住地不變")
                            return (BrHas("辰") && BrHas("戌")) || (BrHas("丑") && BrHas("未"));
                        return false; // 未對應條件，不輸出
                    }

                    case "InjuryInfo":
                        if (cond == "日柱庚午時柱辛巳多心血之疾")
                            return dPillarZ == "庚午" && hPillarZ == "辛巳";
                        if (cond == "日柱乙酉時柱甲申小兒時有肝風之疾")
                            return dPillarZ == "乙酉" && hPillarZ == "甲申";
                        if (cond == "年柱戊申日時乙酉先官司後投井死")
                            return yPillarZ == "戊申" && (dPillarZ == "乙酉" || hPillarZ == "乙酉");
                        if (cond == "年柱戊辰日時癸酉因失盜破產")
                            return yPillarZ == "戊辰" && (dPillarZ == "癸酉" || hPillarZ == "癸酉");
                        if (cond == "年柱甲寅日時柱辛丑定有官刑之災")
                            return yPillarZ == "甲寅" && (dPillarZ == "辛丑" || hPillarZ == "辛丑");
                        if (cond == "四柱寅巳申三刑俱全定有官司或牢獄")
                            return BrHas("寅") && BrHas("巳") && BrHas("申");
                        return true;

                    case "TenGodInfo":
                    {
                        var aSS = new[] { yStemSS, mStemSS, hStemSS, yBranchSS, mBranchSS, dBranchSS, hBranchSS };
                        bool SsGrp(string ss, string g) => g switch {
                            "印"   => ss == "正印" || ss == "偏印",
                            "官殺" => ss == "正官" || ss == "七殺",
                            "比劫" => ss == "比肩" || ss == "劫財",
                            "財"   => ss == "正財" || ss == "偏財",
                            "食傷" => ss == "食神" || ss == "傷官",
                            _ => false
                        };
                        int CntG(string g) => aSS.Count(s => SsGrp(s, g));
                        bool TM(string g) => CntG(g) >= 2;
                        bool Sc(string g) => CntG(g) == 0;
                        if (cond == "印星太多依靠性大無大志")         return TM("印");
                        if (cond == "官殺太多精神萎靡膽小怕事")       return TM("官殺");
                        if (cond == "比劫多不聚財好惹事非")           return TM("比劫");
                        if (cond == "財星太多懼內因財遭災")           return TM("財");
                        if (cond == "食傷太多言語多嘴雜鄙視他人")     return TM("食傷");
                        if (cond == "缺少印星與母親長輩緣薄無靠山")   return Sc("印");
                        if (cond == "缺少比劫人多孤獨靠技藝維生")     return Sc("比劫");
                        if (cond == "缺少食傷行事有恒心善守秘密")     return Sc("食傷");
                        if (cond == "缺少財星財來財去與父親妻子緣薄") return Sc("財");
                        if (cond == "缺少官星不喜拘束女命夫緣薄")     return Sc("官殺");
                        return true;
                    }

                    case "ParentInfo":
                    {
                        var aStemSS = new[] { yStemSS, mStemSS, hStemSS };
                        var aBrSS   = new[] { yBranchSS, mBranchSS, dBranchSS, hBranchSS };
                        bool SsGrp(string ss, string g) => g switch {
                            "印"   => ss == "正印" || ss == "偏印",
                            "比劫" => ss == "比肩" || ss == "劫財",
                            "財"   => ss == "正財" || ss == "偏財",
                            _ => false
                        };
                        bool IsYang(string s) => "甲丙戊庚壬".Contains(s);
                        bool IsYin(string s)  => "乙丁己辛癸".Contains(s);
                        int CntG(string g) => aStemSS.Concat(aBrSS).Count(s => SsGrp(s, g));
                        bool Sc(string g)  => CntG(g) == 0;
                        if (cond == "年干比劫財星論父母")
                            return SsGrp(yStemSS, "比劫") || SsGrp(yBranchSS, "財");
                        if (cond == "正偏印同透有繼母")
                            return aStemSS.Contains("正印") && aStemSS.Contains("偏印");
                        if (cond == "正印透干偏財藏偷生")
                            return aStemSS.Contains("正印") && aBrSS.Contains("偏財");
                        if (cond == "壬乙組合母親為偏房")
                            return (yStem == "壬" && hStem == "乙") || (yStem == "乙" && hStem == "壬");
                        if (cond == "四柱純陽印衰母早喪")
                            return IsYang(yStem) && IsYang(mStem) && IsYang(dStem) && IsYang(hStem) && Sc("印");
                        if (cond == "四柱純陰財衰父早喪")
                            return IsYin(yStem) && IsYin(mStem) && IsYin(dStem) && IsYin(hStem) && Sc("財");
                        if (cond == "年干傷官祖業漂零父母貧困")
                            return yStemSS == "傷官";
                        if (cond == "年干偏財坐驛馬父遠方創業")
                        {
                            // 驛馬以日支三合局定位：申子辰→寅, 寅午戌→申, 亥卯未→巳, 巳酉丑→亥
                            string yiMa = new[]{"申","子","辰"}.Contains(dBranch) ? "寅" :
                                          new[]{"寅","午","戌"}.Contains(dBranch) ? "申" :
                                          new[]{"亥","卯","未"}.Contains(dBranch) ? "巳" :
                                          new[]{"巳","酉","丑"}.Contains(dBranch) ? "亥" : "";
                            return yStemSS == "偏財" && yBranch == yiMa;
                        }
                        if (cond == "年支戌亥印星母有宗教信仰")
                            return (yBranch == "戌" || yBranch == "亥") && SsGrp(yBranchSS, "印");
                        if (cond == "年干支印星喜用書香門第")
                            return SsGrp(yStemSS, "印") || SsGrp(yBranchSS, "印");
                        if (cond == "地支兩見殺為養子")
                        {
                            // 年支三合組 → 殺位：申子辰→戌, 巳酉丑→未, 寅午戌→辰, 亥卯未→丑
                            var oBr3 = new[] { mBranch, dBranch, hBranch };
                            string killBr = new[]{"申","子","辰"}.Contains(yBranch) ? "戌" :
                                            new[]{"巳","酉","丑"}.Contains(yBranch) ? "未" :
                                            new[]{"寅","午","戌"}.Contains(yBranch) ? "辰" :
                                            new[]{"亥","卯","未"}.Contains(yBranch) ? "丑" : "";
                            return !string.IsNullOrEmpty(killBr) && oBr3.Count(b => b == killBr) >= 2;
                        }
                        if (cond == "年干支臨將星父母有權威")
                            // 將星 = 四正地支（三合局中氣）：申子辰→子, 巳酉丑→酉, 寅午戌→午, 亥卯未→卯
                            return new[]{"子","午","卯","酉"}.Contains(yBranch);
                        if (cond == "三柱納音克胎納音父母雙亡" || cond == "年月日時胎支皆克干父母早喪")
                        {
                            // 胎元：天干+1, 地支+3
                            var s10 = new[]{"甲","乙","丙","丁","戊","己","庚","辛","壬","癸"};
                            var b12 = new[]{"子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥"};
                            int mSI = Array.IndexOf(s10, mStem);
                            int mBI = Array.IndexOf(b12, mBranch);
                            if (mSI < 0 || mBI < 0) return false;
                            string tyStem   = s10[(mSI + 1) % 10];
                            string tyBranch = b12[(mBI + 3) % 12];

                            if (cond == "三柱納音克胎納音父母雙亡")
                            {
                                // 六十甲子納音對照表（取末字為五行）
                                var ny60 = new Dictionary<string,string>{
                                    {"甲子","海中金"},{"乙丑","海中金"},{"丙寅","爐中火"},{"丁卯","爐中火"},
                                    {"戊辰","大林木"},{"己巳","大林木"},{"庚午","路旁土"},{"辛未","路旁土"},
                                    {"壬申","劍鋒金"},{"癸酉","劍鋒金"},{"甲戌","山頭火"},{"乙亥","山頭火"},
                                    {"丙子","澗下水"},{"丁丑","澗下水"},{"戊寅","城頭土"},{"己卯","城頭土"},
                                    {"庚辰","白蠟金"},{"辛巳","白蠟金"},{"壬午","楊柳木"},{"癸未","楊柳木"},
                                    {"甲申","泉中水"},{"乙酉","泉中水"},{"丙戌","屋上土"},{"丁亥","屋上土"},
                                    {"戊子","霹靂火"},{"己丑","霹靂火"},{"庚寅","松柏木"},{"辛卯","松柏木"},
                                    {"壬辰","長流水"},{"癸巳","長流水"},{"甲午","沙中金"},{"乙未","沙中金"},
                                    {"丙申","山下火"},{"丁酉","山下火"},{"戊戌","平地木"},{"己亥","平地木"},
                                    {"庚子","壁上土"},{"辛丑","壁上土"},{"壬寅","金箔金"},{"癸卯","金箔金"},
                                    {"甲辰","覆燈火"},{"乙巳","覆燈火"},{"丙午","天河水"},{"丁未","天河水"},
                                    {"戊申","大驛土"},{"己酉","大驛土"},{"庚戌","釵釧金"},{"辛亥","釵釧金"},
                                    {"壬子","桑柘木"},{"癸丑","桑柘木"},{"甲寅","大溪水"},{"乙卯","大溪水"},
                                    {"丙辰","沙中土"},{"丁巳","沙中土"},{"戊午","天上火"},{"己未","天上火"},
                                    {"庚申","石榴木"},{"辛酉","石榴木"},{"壬戌","大海水"},{"癸亥","大海水"}
                                };
                                string tyGz = tyStem + tyBranch;
                                if (!ny60.TryGetValue(tyGz, out string? tyNy) || string.IsNullOrEmpty(tyNy)) return false;
                                char taiE = tyNy[tyNy.Length - 1]; // 胎元納音五行（末字）
                                // 克序：木克土, 土克水, 水克火, 火克金, 金克木
                                var keC = new Dictionary<char,char>{{'木','土'},{'土','水'},{'水','火'},{'火','金'},{'金','木'}};
                                // 統計四柱納音中「克胎元」的數量
                                var pillNy = new[] { yNaYin, mNaYin, dNaYin, hNaYin };
                                int cnt = pillNy.Count(ny =>
                                    !string.IsNullOrEmpty(ny) &&
                                    keC.TryGetValue(ny[ny.Length - 1], out char tgt) && tgt == taiE);
                                return cnt >= 3;
                            }
                            else // 年月日時胎支皆克干父母早喪
                            {
                                // 地支本氣五行
                                var bEMap = new Dictionary<string,string>{
                                    {"子","水"},{"丑","土"},{"寅","木"},{"卯","木"},{"辰","土"},{"巳","火"},
                                    {"午","火"},{"未","土"},{"申","金"},{"酉","金"},{"戌","土"},{"亥","水"}
                                };
                                var keA = new Dictionary<string,string>{{"木","土"},{"土","水"},{"水","火"},{"火","金"},{"金","木"}};
                                // 支克干：支本氣五行克天干五行
                                bool BC(string br, string st) =>
                                    bEMap.TryGetValue(br, out string? bE) &&
                                    keA.TryGetValue(bE, out string? tgt) &&
                                    tgt == KbStemToElement(st);
                                // 五柱（含胎元）皆需支克干
                                return BC(yBranch,yStem) && BC(mBranch,mStem) && BC(dBranch,dStem)
                                    && BC(hBranch,hStem) && BC(tyBranch,tyStem);
                            }
                        }
                        return true;
                    }

                    case "ChildInfo":
                    {
                        var sixComboC = new HashSet<(string, string)>
                        {
                            ("子","丑"),("丑","子"),("寅","亥"),("亥","寅"),
                            ("卯","戌"),("戌","卯"),("辰","酉"),("酉","辰"),
                            ("巳","申"),("申","巳"),("午","未"),("未","午")
                        };
                        var sanXingC = new HashSet<(string,string)>
                        {
                            ("丑","戌"),("戌","丑"),("丑","未"),("未","丑"),("戌","未"),("未","戌"),
                            ("子","卯"),("卯","子"),
                            ("寅","巳"),("巳","寅"),("申","寅"),("寅","申"),("申","巳"),("巳","申")
                        };
                        var aStemSS = new[] { yStemSS, mStemSS, hStemSS };
                        var aBrSS   = new[] { yBranchSS, mBranchSS, dBranchSS, hBranchSS };
                        bool SsGrp(string ss, string g) => g switch {
                            "印"   => ss == "正印" || ss == "偏印",
                            "食傷" => ss == "食神" || ss == "傷官",
                            "官殺" => ss == "正官" || ss == "七殺",
                            "財"   => ss == "正財" || ss == "偏財",
                            _ => false
                        };
                        int CntG(string g) => aStemSS.Concat(aBrSS).Count(s => SsGrp(s, g));
                        bool TM(string g) => CntG(g) >= 2;
                        if (cond == "女命時柱坐梟印克子女")
                            return gender == 2 && (hStemSS == "偏印" || hBranchSS == "偏印");
                        if (cond == "女命日旺時支遇刃梟難產")
                            return gender == 2 && bodyPct > 50 && (hBranchSS == "劫財" || hBranchSS == "偏印");
                        if (cond == "日時相刑女命克夫克子")
                            return gender == 2 && sanXingC.Contains((dBranch, hBranch));
                        if (cond == "日時辰戌相沖中老年克子")
                            return (dBranch == "辰" && hBranch == "戌") || (dBranch == "戌" && hBranch == "辰");
                        if (cond == "時上坐梟年月透財女人有子不死也傷")
                            return gender == 2 && hBranchSS == "偏印" &&
                                   (SsGrp(yStemSS, "財") || SsGrp(mStemSS, "財"));
                        if (cond == "日時相沖中晚年喪子之憂")
                            return clashPairs.Contains((dBranch, hBranch));
                        if (cond == "時帶傷官男命克子")
                            return gender == 1 && hStemSS == "傷官";
                        if (cond == "時干殺旺無制子女不孝叛逆")
                            return hStemSS == "七殺";
                        if (cond == "男命食傷多子女難成氣候")
                            return gender == 1 && TM("食傷");
                        if (cond == "女命印梟多子女難有大發展")
                            return gender == 2 && TM("印");
                        if (cond == "男命時干財官有氣子女有出息")
                            return gender == 1 && (SsGrp(hStemSS, "財") || SsGrp(hStemSS, "官殺"));
                        if (cond == "女命時干財食傷有氣子女有出息")
                            return gender == 2 && (SsGrp(hStemSS, "財") || SsGrp(hStemSS, "食傷"));
                        if (cond == "女命時逢沐浴第一胎難養")
                        {
                            // 沐浴位：依日主陽順陰逆, 甲→子,乙→巳,丙→卯,丁→申,戊→卯,己→申,庚→午,辛→亥,壬→酉,癸→寅
                            var muYuMap = new Dictionary<string,string>{
                                {"甲","子"},{"乙","巳"},{"丙","卯"},{"丁","申"},
                                {"戊","卯"},{"己","申"},{"庚","午"},{"辛","亥"},{"壬","酉"},{"癸","寅"}
                            };
                            return gender == 2 && muYuMap.TryGetValue(dStem, out var muYuBr) && hBranch == muYuBr;
                        }
                        if (cond == "時帶官符出生時父親有官司")
                        {
                            // 口訣: 取太歲前五辰（年支順數+5位）,時支遇之即官符
                            var br12 = new[]{"子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥"};
                            int yIdx = Array.IndexOf(br12, yBranch);
                            if (yIdx >= 0)
                            {
                                string guanFu = br12[(yIdx + 5) % 12];
                                return hBranch == guanFu;
                            }
                            return false;
                        }
                        // 通論: 男命官殺為子女星女命食傷為子女星, 子女星旺衰 - 全部輸出
                        return true;
                    }

                    case "MarriageInfo":
                    {
                        var sixComboM = new HashSet<(string, string)>
                        {
                            ("子","丑"),("丑","子"),("寅","亥"),("亥","寅"),
                            ("卯","戌"),("戌","卯"),("辰","酉"),("酉","辰"),
                            ("巳","申"),("申","巳"),("午","未"),("未","午")
                        };
                        var aStemSS = new[] { yStemSS, mStemSS, hStemSS };
                        var aBrSS   = new[] { yBranchSS, mBranchSS, dBranchSS, hBranchSS };
                        bool SsGrp(string ss, string g) => g switch {
                            "印"   => ss == "正印" || ss == "偏印",
                            "官殺" => ss == "正官" || ss == "七殺",
                            "比劫" => ss == "比肩" || ss == "劫財",
                            "財"   => ss == "正財" || ss == "偏財",
                            "食傷" => ss == "食神" || ss == "傷官",
                            _ => false
                        };
                        int CntG(string g) => aStemSS.Concat(aBrSS).Count(s => SsGrp(s, g));
                        bool TM(string g) => CntG(g) >= 2;
                        bool Sc(string g) => CntG(g) == 0;
                        var otherBrM = new[] { yBranch, mBranch, hBranch };
                        bool dClashed  = otherBrM.Any(b => clashPairs.Contains((dBranch, b)));
                        bool dCombined = otherBrM.Any(b => sixComboM.Contains((dBranch, b)));
                        if (cond == "月柱干傷官支坐七殺女多婚")
                            return gender == 2 && mStemSS == "傷官" && mBranchSS == "七殺";
                        if (cond == "女命日坐傷官必克夫")
                            return gender == 2 && dBranchSS == "傷官";
                        if (cond == "男命坐比劫必妨妻婚姻不順")
                            return gender == 1 && SsGrp(dBranchSS, "比劫");
                        if (cond == "女命甲寅戊申日柱夫有橫死之災")
                            return gender == 2 && (dStem + dBranch == "甲寅" || dStem + dBranch == "戊申");
                        if (cond == "男命日坐偏財主風流偏愛小妾")
                            return gender == 1 && dBranchSS == "偏財";
                        if (cond == "男命日坐印妨妻且妻與母不合")
                            return gender == 1 && SsGrp(dBranchSS, "印");
                        if (cond == "女命官殺均透干多婚外遇")
                            return gender == 2 && aStemSS.Contains("正官") && aStemSS.Contains("七殺");
                        if (cond == "男命干透偏正財多婚外遇")
                            return gender == 1 && aStemSS.Contains("正財") && aStemSS.Contains("偏財");
                        if (cond == "女命傷官旺無財克夫改嫁")
                            return gender == 2 && TM("食傷") && Sc("財");
                        if (cond == "日支被沖夫妻不合難白頭")
                            return dClashed;
                        if (cond == "日支被合化配偶有外遇")
                            return dCombined;
                        if (cond == "男命日支藏財星能得良妻")
                            return gender == 1 && SsGrp(dBranchSS, "財");
                        if (cond == "女命日支藏財星喜用得夫之力")
                            return gender == 2 && SsGrp(dBranchSS, "財");
                        if (cond == "日支子午卯酉配偶漂亮能干")
                            return new[]{"子","午","卯","酉"}.Contains(dBranch);
                        if (cond == "日支寅申巳亥配偶相貌一般聰明能干")
                            return new[]{"寅","申","巳","亥"}.Contains(dBranch);
                        if (cond == "日支辰戌丑未配偶相貌較差樸素")
                            return new[]{"辰","戌","丑","未"}.Contains(dBranch);
                        if (cond == "日支與月支相同配偶漂亮能力強")
                            return dBranch == mBranch;
                        if (cond == "十神傾向食傷生財配偶比自己小")
                            return CntG("食傷") + CntG("財") > CntG("官殺") + CntG("印");
                        if (cond == "十神傾向官印相生配偶比自己大")
                            return CntG("官殺") + CntG("印") > CntG("食傷") + CntG("財");
                        if (cond == "男命財多且無官傷官透有外遇")
                            return gender == 1 && TM("財") && Sc("官殺") && aStemSS.Any(s => SsGrp(s, "食傷"));
                        // 夫妻星喜忌判斷：男以財星, 女以官星
                        {
                            // 財星元素（日主所克）: 木→土, 火→金, 土→水, 金→木, 水→火
                            var keElem = new Dictionary<string,string>{{"木","土"},{"火","金"},{"土","水"},{"金","木"},{"水","火"}};
                            // 官星元素（克日主）: 木←金, 火←水, 土←木, 金←火, 水←土
                            var beKeElem = new Dictionary<string,string>{{"木","金"},{"火","水"},{"土","木"},{"金","火"},{"水","土"}};
                            string marriageElem = gender == 1
                                ? (keElem.TryGetValue(dmElem, out var me1) ? me1 : "")
                                : (beKeElem.TryGetValue(dmElem, out var me2) ? me2 : "");
                            bool marriageStar喜用 = !string.IsNullOrEmpty(marriageElem) && marriageElem == yongShenElem;
                            // 夫妻星坐日支: 日支藏干為夫妻星類型
                            bool marriageOnDay = gender == 1
                                ? (dBranchSS == "正財" || dBranchSS == "偏財")
                                : (dBranchSS == "正官" || dBranchSS == "七殺");
                            if (cond == "夫妻星喜用且坐日支最得力")
                                return marriageStar喜用 && marriageOnDay;
                            if (cond == "夫妻星喜用但日支忌神得力不長久")
                                return marriageStar喜用 && !marriageOnDay;
                        }
                        // 夫妻星遠近看法/結婚離婚標誌: 通論全部輸出
                        return true;
                    }

                    case "BodyTrait":
                    {
                        var aStemSS = new[] { yStemSS, mStemSS, hStemSS };
                        var aBrSS   = new[] { yBranchSS, mBranchSS, dBranchSS, hBranchSS };
                        int CntFood() => aStemSS.Concat(aBrSS).Count(s => s == "食神" || s == "傷官");
                        if (cond == "甲木個子高直挺乙木個高苗條")
                            return dStem == "甲" || dStem == "乙";
                        if (cond == "丙丁火個子可高面紅潤俊俏")
                            return dStem == "丙" || dStem == "丁";
                        if (cond == "戊己土個子矮面色黃")
                            return dStem == "戊" || dStem == "己";
                        if (cond == "庚辛金個子高皮色白")
                            return dStem == "庚" || dStem == "辛";
                        if (cond == "壬癸水個子中等水靈")
                            return dStem == "壬" || dStem == "癸";
                        if (cond == "食傷旺能生財者易發胖")
                            return CntFood() >= 2;
                        if (cond == "八字水木多頭髮烏黑柔細")
                            return (wuXing.TryGetValue("水", out double wv) && wv >= 25) ||
                                   (wuXing.TryGetValue("木", out double muv) && muv >= 25);
                        // 五行代表身體部位/四柱各柱代表身體部位/五行個頭參考值/痣的顏色: 通論全部輸出
                        return true;
                    }

                    default:
                        return true;
                }
            }

            string LfWxSS2(string e) => $"{e}{wuXing[e]:F0}%({LfElemSsGroup(e, dmElem)})";
            string wx = $"{LfWxSS2("木")} {LfWxSS2("火")} {LfWxSS2("土")} {LfWxSS2("金")} {LfWxSS2("水")}";

            // === 表頭 ===
            sb.AppendLine("=================================================================");
            sb.AppendLine("                   玉 洞 子 傳 家 寶 典");
            sb.AppendLine("=================================================================");
            sb.AppendLine($"性別：{genderText}  出生年：{birthYear} 年  虛齡：{currentAge} 歲{curCycleNote}");
            sb.AppendLine($"四柱：{yStem}{yBranch} {mStem}{mBranch} {dStem}{dBranch} {hStem}{hBranch}");
            sb.AppendLine();
            sb.AppendLine("  時辰恐有錯  陰騭最難憑");
            sb.AppendLine("  萬般皆是命  半點不求人");
            sb.AppendLine("-----------------------------------------------------------------");
            sb.AppendLine();

            // 人生指南目錄
            sb.AppendLine("                       人  生  指  南");
            sb.AppendLine("-----------------------------------------------------------------");
            sb.AppendLine("  審時聞切・四時定數");
            sb.AppendLine("  先天八字依古制定");
            sb.AppendLine("  命格判定");
            sb.AppendLine("  用神喜忌");
            sb.AppendLine("  紫微星格");
            sb.AppendLine("  宮星化象（十二宮）");
            sb.AppendLine("  事業格局鑑定");
            sb.AppendLine("  六親緣分鑑定");
            sb.AppendLine("  婚姻深度鑑定");
            sb.AppendLine("  疾厄壽元鑑定");
            sb.AppendLine("  大運逐運論斷");
            sb.AppendLine("  開運指南");
            sb.AppendLine("  出生環境・先天地理風水");
            sb.AppendLine("  人生警示・趨吉避凶");
            sb.AppendLine("  一生命運總評");
            sb.AppendLine("-----------------------------------------------------------------");
            sb.AppendLine();

            // === Ch.1 審時聞切（用獨立 StringBuilder，最後 TrimEnd 避免空白頁）===
            var ch1Sb = new StringBuilder();
            ch1Sb.AppendLine("【第一章：審時聞切 · 四時定數】");
            ch1Sb.AppendLine();
            ch1Sb.Append(LfShiWenSection(hBranch, birthHour, birthMinute, gender));
            ch1Sb.Append(LfBaiShengSections(yBranch, hBranch, birthHour, birthMinute, birthYear, lunarMonth, gender, mBranch, dStem));

            // 中原盲派 - 時支/時干直斷
            if (zRules.Count > 0)
            {
                string hBranchGroup = hBranch is "子" or "午" or "卯" or "酉" ? "子午卯酉"
                    : hBranch is "寅" or "申" or "巳" or "亥" ? "寅申巳亥"
                    : hBranch is "辰" or "戌" or "丑" or "未" ? "辰戌丑未" : "";
                string hStemGroup = hStem is "甲" or "乙" ? "甲乙時干"
                    : hStem is "丙" or "丁" ? "丙丁時干"
                    : hStem is "戊" or "己" ? "戊己時干"
                    : hStem is "庚" or "辛" ? "庚辛時干"
                    : hStem is "壬" or "癸" ? "壬癸時干" : "";
                var hBranchRule = zRules.FirstOrDefault(r => r.RuleType == "HourBranch" && r.Condition == hBranchGroup);
                var hStemRule   = zRules.FirstOrDefault(r => r.RuleType == "HourStem"   && r.Condition == hStemGroup);
                if (hBranchRule != null || hStemRule != null)
                {
                    ch1Sb.AppendLine("【時柱印記】");
                    if (hBranchRule != null) { ch1Sb.AppendLine($"▍時支（{hBranchGroup}）"); ch1Sb.AppendLine(hBranchRule.Content.TrimEnd()); }
                    if (hStemRule   != null) { ch1Sb.AppendLine($"▍時干（{hStem}）"); ch1Sb.AppendLine(hStemRule.Content.TrimEnd()); }
                }
            }
            // 去除第一章尾部所有空白，確保第二章前不產生空白頁
            sb.AppendLine(ch1Sb.ToString().TrimEnd());

            // === Ch.2 四柱根苗花果 ===
            sb.AppendLine("【第二章：先天八字依古制定】");
            sb.AppendLine();
            sb.AppendLine("一、根苗花果");
            sb.AppendLine("| 項目 | 時柱 | 日柱 | 月柱 | 年柱 |");
            sb.AppendLine("|------|------|------|------|------|");
            sb.AppendLine($"| 六神 | {hStemSS} | 元神 | {mStemSS} | {yStemSS} |");
            sb.AppendLine($"| 天干 | {hStem} | {dStem} | {mStem} | {yStem} |");
            sb.AppendLine($"| 地支 | {hBranch} | {dBranch} | {mBranch} | {yBranch} |");
            sb.AppendLine($"| 藏神 | {LfFmtHidden(hBranch,dStem)} | {LfFmtHidden(dBranch,dStem)} | {LfFmtHidden(mBranch,dStem)} | {LfFmtHidden(yBranch,dStem)} |");
            sb.AppendLine($"| 納音 | {hNaYin} | {dNaYin} | {mNaYin} | {yNaYin} |");
            var (wang, xiang, xiu, qiu, si) = LfGetWangXiang(mBranch);
            sb.AppendLine($"| 旺相 | {wang}旺 | {xiang}相 | {xiu}休 | {qiu}囚 {si}死 |");
            sb.AppendLine();
            sb.AppendLine("二、天干十神");
            string[] allStems10 = { "甲","乙","丙","丁","戊","己","庚","辛","壬","癸" };
            sb.AppendLine("| 天干 | " + string.Join(" | ", allStems10) + " |");
            sb.AppendLine("|------|" + string.Join("|", allStems10.Select(_ => "----")) + "|");
            sb.AppendLine("| 十神 | " + string.Join(" | ", allStems10.Select(s => LfShiShenAbbr(s, dStem).PadLeft(2))) + " |");
            sb.AppendLine();
            sb.AppendLine("三、地支藏神十神");
            string[] allBrs12 = { "子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥" };
            // 橫排：地支為欄
            sb.AppendLine("| 項目 | " + string.Join(" | ", allBrs12.Select(br => br + (branches.Contains(br) ? "★" : ""))) + " |");
            sb.AppendLine("|------" + string.Join("|", allBrs12.Select(_ => "----")) + "|");
            sb.AppendLine("| 藏神 | " + string.Join(" | ", allBrs12.Select(br => LfFmtHidden(br, dStem))) + " |");
            double biJiPct = wuXing.GetValueOrDefault(dmElem, 0) + wuXing.GetValueOrDefault(LfGenByElem.GetValueOrDefault(dmElem, ""), 0);
            sb.AppendLine($"比印陣計分：{biJiPct:F0}%");
            sb.AppendLine();

            // === Ch.3 深度論斷 ===
            sb.AppendLine($"【第三章：日柱深度論斷 · {dStem}{dBranch}】");
            sb.AppendLine();
            if (kb == null)
            {
                sb.AppendLine("（此日柱斷語尚待補充）");
            }
            else
            {
                void AppKb(string label, string? val)
                {
                    if (!string.IsNullOrWhiteSpace(val)) { sb.AppendLine($"▍{label}"); sb.AppendLine(val); sb.AppendLine(); }
                }
                AppKb("核心", kb.Overview);
                AppKb("神殺特質", kb.ShenAnalysis);
                AppKb("內在特質", kb.InnerTraits);
                AppKb("事業傾向", kb.Career);
                AppKb("天生弱點", kb.Weaknesses);
                string kbSeasonChar = "寅卯辰".Contains(mBranch) ? "春"
                    : "巳午未".Contains(mBranch) ? "夏"
                    : "申酉戌".Contains(mBranch) ? "秋" : "冬";
                AppKb("月令影響", LfFilterSeasonText(kb.MonthInfluence, kbSeasonChar));
                AppKb(gender == 1 ? "男命論斷" : "女命論斷", LfFilterSeasonText(gender == 1 ? kb.MaleChart : kb.FemaleChart, kbSeasonChar));
                AppKb("最佳時辰", kb.SpecialHours);
            }
            // === 十干象法 ===
            string shiGanXiangDesc = LfShiGanXiangFa(dStem, mBranch);
            if (!string.IsNullOrEmpty(shiGanXiangDesc))
            {
                sb.AppendLine("【十干象法】");
                sb.AppendLine(shiGanXiangDesc);
                sb.AppendLine();
            }

            // === 納音論斷 ===
            string nayinDesc = LfNaYin(yStem, yBranch, mBranch, dStem, dBranch, hBranch);
            if (!string.IsNullOrEmpty(nayinDesc))
            {
                sb.AppendLine("【納音論斷】");
                sb.AppendLine(nayinDesc);
                sb.AppendLine();
            }

            // === 空亡論斷 ===
            string kongWangDesc = LfKongWang(yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch);
            if (!string.IsNullOrEmpty(kongWangDesc))
            {
                sb.AppendLine("【空亡論斷】");
                sb.AppendLine(kongWangDesc);
                sb.AppendLine();
            }

            // === 四柱神煞 ===
            string shenShaDesc = LfShenSha(
                yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch,
                yStemSS, mStemSS, hStemSS,
                yBranchSS, mBranchSS, dBranchSS, hBranchSS);
            if (!string.IsNullOrEmpty(shenShaDesc))
            {
                sb.AppendLine("【四柱神煞】");
                sb.AppendLine(shenShaDesc);
                sb.AppendLine();
            }

            // 中原盲派 - 天干地支重複直斷
            if (zRules.Count > 0 && (repeatedStems.Count > 0 || repeatedBranches.Count > 0))
            {
                sb.AppendLine("【四柱干支特徵論斷】");
                sb.AppendLine();
                foreach (var stem in repeatedStems)
                {
                    var matchKey = $"三{stem}";
                    var rules = zRules.Where(r => r.RuleType == "StemRepeat" && r.Condition.Contains(stem)).ToList();
                    if (rules.Count > 0)
                    {
                        sb.AppendLine($"▍天干 {stem} 出現三次以上：");
                        foreach (var r in rules) sb.AppendLine($"• {r.Content}");
                        sb.AppendLine();
                    }
                }
                foreach (var branch in repeatedBranches)
                {
                    var rules = zRules.Where(r => r.RuleType == "BranchRepeat" && r.Condition.Contains(branch)).ToList();
                    if (rules.Count > 0)
                    {
                        sb.AppendLine($"▍地支 {branch} 出現三次以上：");
                        foreach (var r in rules) sb.AppendLine($"• {r.Content}");
                        sb.AppendLine();
                    }
                }
            }

            // === Ch.4 命局格局判定 ===
            sb.AppendLine("【第四章：命格判定】");
            sb.AppendLine();
            sb.AppendLine("【命局體性（寒暖濕燥）】");
            sb.AppendLine($"月支 {mBranch} 生人，命局屬【{seaLabel}】。");
            if (seaLabel == "寒凍")
                sb.AppendLine("最喜：丙丁巳午火暖局。最忌：壬癸亥子水助寒。調候急用：丙丁火。");
            else if (seaLabel == "炎熱")
                sb.AppendLine("最喜：壬癸亥子水消暑，庚辛申酉金。最忌：丙丁巳午火助熱。調候急用：壬癸水。");
            else
                sb.AppendLine("體性溫和，以日主強弱論用神，無需特別調候。");
            sb.AppendLine();
            sb.AppendLine("【日主強弱判定】");
            sb.AppendLine($"日干 {dStem}（{dmElem}），月令 {mBranch}（{season}季）。");
            sb.AppendLine($"五行分布：{wx}");
            double tiPct2 = wuXing.GetValueOrDefault(dmElem, 0) + wuXing.GetValueOrDefault(LfGenByElem.GetValueOrDefault(dmElem, ""), 0) + wuXing.GetValueOrDefault(LfElemGen.GetValueOrDefault(dmElem, ""), 0);
            double yongPct2 = wuXing.GetValueOrDefault(LfElemOvercome.GetValueOrDefault(dmElem, ""), 0) + wuXing.GetValueOrDefault(LfElemOvercomeBy.GetValueOrDefault(dmElem, ""), 0);
            sb.AppendLine($"比印陣：{biJiPct:F0}% | 洩克陣：{100 - biJiPct:F0}%   (印比食)體 {tiPct2:F0}%  (財官)用 {yongPct2:F0}%");
            sb.AppendLine($"結論：日主【{bodyLabel}】（強弱度：{bodyPct:F0}%）");
            // 節氣司令用事（驗算日主強弱參考）
            if (birthMonth.HasValue && birthDay.HasValue)
            {
                string siLingDesc = LfCalcSiLing(birthYear, birthMonth.Value, birthDay.Value, mBranch, dStem, calDb);
                if (!string.IsNullOrEmpty(siLingDesc))
                {
                    sb.AppendLine();
                    sb.AppendLine("【節氣司令用事（驗算強弱參考）】");
                    sb.AppendLine(siLingDesc);
                }
            }
            sb.AppendLine();
            sb.AppendLine("【格局判定】");
            sb.AppendLine($"格局：【{pattern}】");
            sb.AppendLine(LfPatternDesc(pattern));
            // 四柱神殺
            var shenShaList = LfGetBaziShenSha(yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch);
            if (shenShaList.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("【星曜影響】");
                foreach (var ss in shenShaList)
                    sb.AppendLine(ss);
            }
            sb.AppendLine();

            // === Ch.5 格局與用神判定 ===
            sb.AppendLine("【第五章：用神喜忌】");
            sb.AppendLine();
            sb.AppendLine($"用神：【{yongShenElem}】（理由：{yongReason}）");
            sb.AppendLine($"喜用：天干 {LfElemStems(yongShenElem)}，地支 {LfElemBranches(yongShenElem)}");
            if (fuYiElem != yongShenElem)
                sb.AppendLine($"輔助喜神：【{fuYiElem}】（{(bodyPct <= 40 ? "印比互補扶身" : "官財互補制衡")}）");
            string tuneElemV2 = season == "冬" ? "火" : season == "夏" ? "水" : "";
            if (!string.IsNullOrEmpty(tuneElemV2) && tuneElemV2 != yongShenElem && tuneElemV2 != fuYiElem)
                sb.AppendLine($"調候喜神：【{tuneElemV2}】（{(season == "冬" ? "冬月寒凍，喜火暖局" : "夏月炎熱，喜水消暑")}）");
            sb.AppendLine($"大忌(X)：{jiShenElem}，天干 {LfElemStems(jiShenElem)}，地支 {LfElemBranches(jiShenElem)}");
            string jiYongElemV2 = LfElemOvercomeBy.GetValueOrDefault(yongShenElem, "");
            if (!string.IsNullOrEmpty(jiYongElemV2) && jiYongElemV2 != jiShenElem)
                sb.AppendLine($"次忌(△忌)：{jiYongElemV2}（克用神 {yongShenElem}）");
            sb.AppendLine();
            sb.AppendLine(LfBuildYongJiTable(yongShenElem, fuYiElem, jiShenElem, tuneElemV2, dStem, branches));
            sb.AppendLine();

            // === Ch.6 紫微格局論 ===
            sb.AppendLine("【第六章：紫微星格】");
            sb.AppendLine();
            if (!hasZiwei)
            {
                sb.AppendLine("（命盤資料不含紫微斗數，略去此章）");
            }
            else
            {
                if (!string.IsNullOrEmpty(wuXingJuText)) sb.AppendLine($"【五行局】{wuXingJuText}　命主星：{mingZhu}　身主星：{shenZhu}");
                sb.AppendLine();
                if (!string.IsNullOrEmpty(nianSiHuaXing))
                    sb.AppendLine($"【先天特性】{StripNianShengRen(nianSiHuaXing)}");
                if (!string.IsNullOrEmpty(ziweiMing))
                    sb.AppendLine($"【主星宮位】{ziweiMing}");
                if (!string.IsNullOrEmpty(starDescMing))
                    sb.AppendLine($"【命宮星情】{starDescMing}");
                if (doubleDescs != null && doubleDescs.TryGetValue("命宮", out var dblMing) && !string.IsNullOrEmpty(dblMing))
                    sb.AppendLine($"【命宮雙星論斷】{dblMing}");
                if (minorDescs != null && minorDescs.TryGetValue("命宮", out var minMing) && !string.IsNullOrEmpty(minMing))
                    sb.AppendLine($"【命宮輔星加臨】{minMing}");
                if (!string.IsNullOrEmpty(siHuaLu))
                    sb.AppendLine($"【先天化祿·{siHuaLuPalace}】{siHuaLu}");
                if (!string.IsNullOrEmpty(siHuaQuan))
                    sb.AppendLine($"【先天化權·{siHuaQuanPalace}】{siHuaQuan}");
                if (!string.IsNullOrEmpty(siHuaKe))
                    sb.AppendLine($"【先天化科·{siHuaKePalace}】{siHuaKe}");
                if (!string.IsNullOrEmpty(siHuaJi))
                    sb.AppendLine($"【先天化忌·{siHuaJiPalace}】{siHuaJi}");
                // 命宮格局論斷（從原第八章移入）
                if (!string.IsNullOrEmpty(ziweiGeJuContent))
                {
                    sb.AppendLine();
                    sb.AppendLine("【命宮格局論斷】");
                    sb.AppendLine(ziweiGeJuContent.TrimEnd());
                }
                // 各宮宮干飛星四化（從原第八、九章移入）
                bool anySiHuaPal = false;
                foreach (var kv in siHua)
                {
                    if (!string.IsNullOrEmpty(kv.Value.txt))
                    {
                        sb.AppendLine(kv.Value.txt);
                        anySiHuaPal = true;
                    }
                }
                if (!anySiHuaPal) sb.AppendLine("（四化 KB 資料待補充）");
            }
            sb.AppendLine();

            // === Ch.7 宮星化象 ===
            sb.AppendLine("【第七章：宮星化象（十二宮）】");
            sb.AppendLine();
            if (!hasZiwei || string.IsNullOrEmpty(ziweiFullContent))
            {
                sb.AppendLine("（紫微 KB 資料不足，略去此章）");
            }
            else
            {
                string[] palaceNamesAll = { "命宮","兄弟宮","夫妻宮","子女宮","財帛宮","疾厄宮","遷移宮","奴僕宮","官祿宮","田宅宮","福德宮","父母宮" };
                string[] sectionKeys   = { "命宮","兄弟宮","夫妻宮","子女宮","財帛宮","疾厄宮","遷移宮","交友宮","事業宮","田宅宮","福德宮","父母宮" };
                string[] palaceLookups = { "命宮","兄弟","夫妻","子女","財帛","疾厄","遷移","奴僕","官祿","田宅","福德","父母" };
                for (int i = 0; i < palaceNamesAll.Length; i++)
                {
                    if (palaceNamesAll[i] == "命宮") continue; // 第六章已描述命宮，此處略去
                    // 以下宮位有專屬章節深度論斷，此處略去避免重複
                    if (palaceNamesAll[i] == "官祿宮" || palaceNamesAll[i] == "財帛宮" ||
                        palaceNamesAll[i] == "父母宮" || palaceNamesAll[i] == "子女宮" ||
                        palaceNamesAll[i] == "夫妻宮" || palaceNamesAll[i] == "疾厄宮") continue;
                    string pStars  = KbGetPalaceStars(palacesYdz, palaceLookups[i]);
                    string pBranch = KbGetPalaceBranch(palacesYdz, palaceLookups[i]);
                    string kbContent = KbFilterZiweiContent(KbExtractPalaceSection(ziweiFullContent, sectionKeys[i]), KbGetPalaceStarsSet(palacesYdz, palaceLookups[i]), chartStars).Trim();
                    sb.AppendLine("────────────────────────────────────");
                    string brLabel = string.IsNullOrEmpty(pBranch) ? "" : $"（坐{pBranch}）";
                    string starsLabel = string.IsNullOrEmpty(pStars) ? "空宮" : pStars;
                    sb.AppendLine($"● {palaceNamesAll[i]}{brLabel} - 【{starsLabel}】");
                    // 星情特質（KbQueryStarInPalace）
                    if (allPalaceStarDescs != null && allPalaceStarDescs.TryGetValue(palaceNamesAll[i], out var palStarDesc) && !string.IsNullOrEmpty(palStarDesc))
                    {
                        // 非疾厄宮：過濾掉含健康/疾病詞彙的行（這類內容只適合在疾厄宮呈現）
                        string filteredStarDesc = palStarDesc;
                        if (palaceNamesAll[i] != "疾厄宮")
                        {
                            var healthKeywords = new[] { "腎臟", "腎虛", "腎", "疾病", "病症", "生病", "患病", "疾厄", "泌尿" };
                            var descLines = palStarDesc.Split('\n');
                            var filteredLines = descLines.Where(l => !healthKeywords.Any(k => l.Contains(k))).ToArray();
                            filteredStarDesc = string.Join("\n", filteredLines).Trim();
                        }
                        if (!string.IsNullOrEmpty(filteredStarDesc))
                            sb.AppendLine($"▶ 星情特質：{filteredStarDesc}");
                    }
                    if (!string.IsNullOrEmpty(kbContent))
                        sb.AppendLine($"▶ 特性診斷：{kbContent}");
                    else
                        sb.AppendLine("▶ 特性診斷：（待補充）");
                    // 空宮：額外顯示對宮主星作為參考
                    if (string.IsNullOrEmpty(pStars))
                    {
                        int opIdx = (i + 6) % 12;
                        string opStars = KbGetPalaceStars(palacesYdz, palaceLookups[opIdx]);
                        if (!string.IsNullOrEmpty(opStars))
                            sb.AppendLine($"▶ 對宮（{palaceNamesAll[opIdx]}）主星：{opStars}");
                    }
                    // 雙星 + 輔星
                    string dblKey = KbNormalizePalaceName(palaceNamesAll[i]);
                    if (doubleDescs != null && doubleDescs.TryGetValue(dblKey, out var dblTxt) && !string.IsNullOrEmpty(dblTxt))
                        sb.AppendLine($"▶ 雙星論斷：{dblTxt}");
                    if (minorDescs != null && minorDescs.TryGetValue(dblKey, out var minTxt) && !string.IsNullOrEmpty(minTxt))
                        sb.AppendLine($"▶ 輔星加臨：{minTxt}");
                }
                sb.AppendLine("────────────────────────────────────");
            }
            sb.AppendLine();

            // === Ch.8 事業格局鑑定 ===
            sb.AppendLine("【第八章：事業格局鑑定】");
            sb.AppendLine();

            // 【事業特質】：五行職業取象 + 官祿宮主星（日柱事業傾向已於第三章論述）
            sb.AppendLine("【事業特質】");
            var (cfHuangliang, cfYangZhiYin) = KbCalcCareerFlags(
                yStemSS, yBranchSS, mStemSS, mBranchSS,
                dBranchSS, hStemSS, hBranchSS, dStem, pattern);
            sb.AppendLine("【五行職業取象（適合行業）】");
            sb.Append(KbSanmenJobByElem(dmElem, pattern, yongShenElem, cfHuangliang, cfYangZhiYin));
            if (!string.IsNullOrEmpty(ziweiOff))
            {
                sb.AppendLine();
                sb.AppendLine($"【官祿宮主星·{offStars}（事業個性）】");
                sb.AppendLine(ziweiOff);
            }
            if (!string.IsNullOrEmpty(starDescOff))
                sb.AppendLine($"【官祿星性】{starDescOff}");
            if (siHua.TryGetValue("官祿化祿", out var offLu) && !string.IsNullOrEmpty(offLu.txt))
                sb.AppendLine($"【官祿化祿飛{offLu.pal}】{offLu.txt}");
            if (siHua.TryGetValue("官祿化忌", out var offJi) && !string.IsNullOrEmpty(offJi.txt))
                sb.AppendLine($"【官祿化忌飛{offJi.pal}】{offJi.txt}");
            if (doubleDescs != null && doubleDescs.TryGetValue("官祿宮", out var dblOff) && !string.IsNullOrEmpty(dblOff))
                sb.AppendLine($"【官祿雙星論斷】{dblOff}");
            if (minorDescs != null && minorDescs.TryGetValue("官祿宮", out var minOff) && !string.IsNullOrEmpty(minOff))
                sb.AppendLine($"【官祿輔星加臨】{minOff}");
            // 財帛宮紫微鑑定
            if (hasZiwei && !string.IsNullOrEmpty(ziweiWlt))
            {
                sb.AppendLine();
                sb.AppendLine($"【財帛宮主星·{wltStars}（財富個性）】");
                sb.AppendLine(ziweiWlt);
            }
            if (!string.IsNullOrEmpty(starDescWlt))
                sb.AppendLine($"【財帛星性】{starDescWlt}");
            if (siHua.TryGetValue("財帛化祿", out var wltLu) && !string.IsNullOrEmpty(wltLu.txt))
                sb.AppendLine($"【財帛化祿飛{wltLu.pal}】{wltLu.txt}");
            if (siHua.TryGetValue("財帛化忌", out var wltJi) && !string.IsNullOrEmpty(wltJi.txt))
                sb.AppendLine($"【財帛化忌飛{wltJi.pal}】{wltJi.txt}");
            if (doubleDescs != null && doubleDescs.TryGetValue("財帛宮", out var dblWlt) && !string.IsNullOrEmpty(dblWlt))
                sb.AppendLine($"【財帛雙星論斷】{dblWlt}");
            if (minorDescs != null && minorDescs.TryGetValue("財帛宮", out var minWlt) && !string.IsNullOrEmpty(minWlt))
                sb.AppendLine($"【財帛輔星加臨】{minWlt}");
            sb.AppendLine();

            // 事業格局判定（過三關分析）
            sb.AppendLine(KbSanmenCareer(
                yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch,
                yStemSS, mStemSS, hStemSS, yBranchSS, mBranchSS, dBranchSS, hBranchSS,
                dmElem, pattern, bodyPct, yongShenElem, jiShenElem, wuXing));
            sb.AppendLine();

            // 中原盲派 - 職業直斷參照
            AppZrList(sb, zRules.Where(r => r.RuleType == "CareerInfo" && ZrApplies(r)), "【職業從業補充論斷】");
            AppZrList(sb, zRules.Where(r => r.RuleType == "TenGodInfo" && ZrApplies(r)), "【十神性質補充論斷】");

            // === Ch.9 六親緣分鑑定 ===
            sb.AppendLine("【第九章：六親緣分鑑定】");
            sb.AppendLine();
            {
                string ageHint = LfAgeTopicHint(currentAge);
                if (!string.IsNullOrEmpty(ageHint)) sb.AppendLine(ageHint);
            }
            // 六親日柱影響（DB 有內容優先，否則依五行百分比自動計算）
            {
                string fatherElem = dmElem switch { "木"=>"土","火"=>"金","土"=>"水","金"=>"木","水"=>"火",_=>"" };
                string siblingElem = dmElem;
                string childElem   = dmElem switch { "木"=>"火","火"=>"土","土"=>"金","金"=>"水","水"=>"木",_=>"" };
                double fPct = wuXing.GetValueOrDefault(fatherElem, 0);
                double sPct = wuXing.GetValueOrDefault(siblingElem, 0);
                double cPct = wuXing.GetValueOrDefault(childElem, 0);

                string fatherDesc = !string.IsNullOrEmpty(kb?.FatherInfluence) ? kb.FatherInfluence :
                    fPct >= 35 ? $"日主（{dmElem}）以{fatherElem}為偏財父星，命局{fatherElem}氣旺（{fPct:F0}%），父親能幹有成，在家中影響力大，對命主期望高，親子關係深厚。" :
                    fPct >= 20 ? $"日主（{dmElem}）以{fatherElem}為偏財父星，命局{fatherElem}氣中等（{fPct:F0}%），父親盡職持家，能給予命主基本支持，親子關係平順。" :
                    fPct >= 10 ? $"日主（{dmElem}）以{fatherElem}為偏財父星，命局{fatherElem}氣偏弱（{fPct:F0}%），父親早年較辛苦，與命主互動有限，命主較早養成自立個性。" :
                               $"日主（{dmElem}）以{fatherElem}為偏財父星，命局{fatherElem}氣極弱（{fPct:F0}%），父緣偏薄，命主與父親聚少離多，靠自身奮鬥成長。";

                string siblingDesc = !string.IsNullOrEmpty(kb?.SiblingInfluence) ? kb.SiblingInfluence :
                    sPct >= 35 ? $"比劫（{siblingElem}）在命局旺盛（{sPct:F0}%），兄弟姐妹緣深，手足情誼濃厚，但競爭意識強，需防因財務資源產生摩擦。" :
                    sPct >= 20 ? $"比劫（{siblingElem}）力量適中（{sPct:F0}%），手足關係平衡，兄弟姐妹感情穩定，彼此尊重各自空間。" :
                    sPct >= 10 ? $"比劫（{siblingElem}）偏弱（{sPct:F0}%），兄弟姐妹人數較少或手足緣薄，凡事較需靠自己解決。" :
                               $"比劫（{siblingElem}）極弱（{sPct:F0}%），手足緣薄，兄弟姐妹分散各地，命主個性獨立自主。";

                string childDesc = !string.IsNullOrEmpty(kb?.ChildInfluence) ? kb.ChildInfluence :
                    cPct >= 35 ? $"食傷（{childElem}）在命局旺盛（{cPct:F0}%），子女緣深，子女聰明有才，親子感情濃厚，但需防過度干涉子女成長。" :
                    cPct >= 20 ? $"食傷（{childElem}）力量適中（{cPct:F0}%），子女緣份平順，子女有一定才幹，親子關係和諧自然。" :
                    cPct >= 10 ? $"食傷（{childElem}）偏弱（{cPct:F0}%），子女緣較薄，或子女個性獨立，需主動用心經營親子感情。" :
                               $"食傷（{childElem}）極弱（{cPct:F0}%），子女緣薄，或子女較少，命主需特別用心才能維持深厚親子情。";

                // 母親
                if (!string.IsNullOrEmpty(kb?.MotherInfluence))
                {
                    sb.AppendLine("【日柱母星影響】");
                    sb.AppendLine(kb.MotherInfluence);
                    sb.AppendLine();
                }
                sb.AppendLine("【日柱父星影響】");
                sb.AppendLine(fatherDesc);
                sb.AppendLine();
                sb.AppendLine("【日柱手足影響】");
                sb.AppendLine(siblingDesc);
                sb.AppendLine();
                sb.AppendLine("【日柱子女影響】");
                sb.AppendLine(childDesc);
                sb.AppendLine();
            }
            sb.AppendLine(KbSanmenSixRelatives(
                yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch,
                yStemSS, mStemSS, hStemSS, yBranchSS, mBranchSS, dBranchSS, hBranchSS,
                dmElem, pattern, bodyPct, yongShenElem, jiShenElem, wuXing, gender, birthYear, scored));
            sb.AppendLine();

            // 紫微父母宮/子女宮補充
            if (hasZiwei)
            {
                if (!LfShouldSkipPalace("父母宮", currentAge) && !string.IsNullOrEmpty(ziweiPar))
                    sb.AppendLine($"【{LfPalaceAgeLabel("父母宮", currentAge)}·{ziweiParStar}】{ziweiPar}");
                if (!LfShouldSkipPalace("子女宮", currentAge) && !string.IsNullOrEmpty(ziweiCld))
                    sb.AppendLine($"【{LfPalaceAgeLabel("子女宮", currentAge)}·{ziweiCldStar}】{ziweiCld}");
            }

            // 中原盲派 - 六親直斷參照
            AppZrList(sb, zRules.Where(r => r.RuleType == "ParentInfo"  && ZrApplies(r)), "【父母緣份補充論斷】");
            AppZrList(sb, zRules.Where(r => r.RuleType == "SiblingInfo" && ZrApplies(r)), "【兄弟姐妹補充論斷】");
            AppZrList(sb, zRules.Where(r => r.RuleType == "ChildInfo"   && ZrApplies(r)), "【子女緣份補充論斷】");

            // === Ch.10 婚姻深度鑑定 ===
            sb.AppendLine("【第十章：婚姻深度鑑定】");
            sb.AppendLine();
            // 紫微夫妻宮補充（日柱男/女命論斷已於第三章論述）
            if (hasZiwei)
            {
                if (!string.IsNullOrEmpty(ziweiSps))
                {
                    // 過濾夫妻宮 KB 中混入的事業宮論述（含「事業」的行，但保留同時含「夫妻」或「配偶」的行）
                    string spsFiltered = KbRemoveOffTopicLines(ziweiSps,
                        new[] { "事業", "轉業", "工藝", "工業設計", "工藝美術", "從商", "職業", "行業" },
                        keepIfContains: new[] { "夫妻", "配偶", "感情" });
                    if (!string.IsNullOrEmpty(spsFiltered))
                        sb.AppendLine($"【夫妻宮主星·{spsStars}（感情個性）】{spsFiltered}");
                }
                if (!string.IsNullOrEmpty(starDescSps))
                    sb.AppendLine($"【夫妻星性】{starDescSps}");
                if (siHua.TryGetValue("夫妻化祿", out var spsLu) && !string.IsNullOrEmpty(spsLu.txt))
                    sb.AppendLine($"【夫妻化祿飛{spsLu.pal}】{spsLu.txt}");
                if (siHua.TryGetValue("夫妻化忌", out var spsJi) && !string.IsNullOrEmpty(spsJi.txt))
                    sb.AppendLine($"【夫妻化忌飛{spsJi.pal}】{spsJi.txt}");
                if (doubleDescs != null && doubleDescs.TryGetValue("夫妻宮", out var dblSps) && !string.IsNullOrEmpty(dblSps))
                    sb.AppendLine($"【夫妻雙星論斷】{dblSps}");
                if (minorDescs != null && minorDescs.TryGetValue("夫妻宮", out var minSps) && !string.IsNullOrEmpty(minSps))
                    sb.AppendLine($"【夫妻輔星加臨】{minSps}");
                if (string.IsNullOrEmpty(ziweiSps)) sb.AppendLine("（紫微夫妻宮資料待補充）");
            }
            sb.AppendLine();

            // 中原盲派 - 婚姻直斷參照
            AppZrList(sb, zRules.Where(r => r.RuleType == "MarriageInfo" && ZrApplies(r)), "【婚姻配偶補充論斷】");

            // === Ch.11 疾厄壽元鑑定 ===
            sb.AppendLine("【第十一章：疾厄壽元鑑定】");
            sb.AppendLine();
            // 日柱天生弱點已於第三章論述
            sb.AppendLine(KbSanmenHealthLongevity(
                yStem, mStem, hStem, yBranch, mBranch, dBranch, hBranch,
                dStem, dmElem, bodyPct, yongShenElem, jiShenElem,
                wuXing, season, seaLabel, scored, currentAge));
            // 紫微疾厄宮補充
            if (hasZiwei)
            {
                if (!string.IsNullOrEmpty(ziweiHlt))
                    sb.AppendLine($"【疾厄宮主星·{hltStars}】{ziweiHlt}");
                if (!string.IsNullOrEmpty(starDescHlt))
                    sb.AppendLine($"【疾厄星性】{starDescHlt}");
                if (siHua.TryGetValue("疾厄化忌", out var hltJi) && !string.IsNullOrEmpty(hltJi.txt))
                    sb.AppendLine($"【疾厄化忌飛{hltJi.pal}】{hltJi.txt}");
                if (doubleDescs != null && doubleDescs.TryGetValue("疾厄宮", out var dblHlt) && !string.IsNullOrEmpty(dblHlt))
                    sb.AppendLine($"【疾厄雙星論斷】{dblHlt}");
                if (minorDescs != null && minorDescs.TryGetValue("疾厄宮", out var minHlt) && !string.IsNullOrEmpty(minHlt))
                    sb.AppendLine($"【疾厄輔星加臨】{minHlt}");
            }
            sb.AppendLine();

            // 中原盲派 - 傷病牢獄直斷參照
            AppZrList(sb, zRules.Where(r => r.RuleType == "InjuryInfo" && ZrApplies(r)), "【傷病牢獄補充論斷】");
            AppZrList(sb, zRules.Where(r => r.RuleType == "BodyTrait"  && ZrApplies(r)), "【身體特徵補充論斷】");

            // === Ch.12 大運逐運論斷 ===
            sb.AppendLine("【第十二章：大運逐運論斷（天干地支分期）】");
            sb.AppendLine();
            if (scored.Count == 0)
            {
                sb.AppendLine("（大運資料不足）");
            }
            else
            {
                // 預先計算旬空（供干支引動用）
                var dayEmpty14 = LfCalcDayEmpty(dStem, dBranch);
                var pillarStems14    = new[] { yStem, mStem, dStem, hStem };
                var pillarBranches14 = new[] { yBranch, mBranch, dBranch, hBranch };
                var pillarBranchSS14 = new[] { yBranchSS, mBranchSS, dBranchSS, hBranchSS };

                bool isBodyStrong14 = bodyPct >= 50;
                var (goodElems14, badElems14) = LfGetPatternLuckElems(pattern, yongShenElem, fuYiElem, dmElem, isBodyStrong14);

                // 只顯示目前大運前一步到往後第8步（過去已過，無論斷意義）
                int curIdx14 = scored.FindIndex(lc => lc.startAge <= currentAge && lc.endAge >= currentAge);
                int startIdx14 = Math.Max(0, curIdx14 - 1);
                var scored14 = scored.Skip(startIdx14).Take(8).ToList();

                // 一、大運走勢總覽表
                sb.AppendLine("一、大運走勢總覽");
                sb.AppendLine($"（顯示範圍：前一步大運至往後第8步，共 {scored14.Count} 步；全部 {scored.Count} 步）");
                sb.AppendLine("| 歲數 | 干支 | 天干 | 地支 | 評分 | 等級 |");
                sb.AppendLine("|------|------|------|------|------|------|");
                foreach (var lc14 in scored14)
                {
                    string bSS14 = LfBranchHiddenRatio.TryGetValue(lc14.branch, out var bhAll14) && bhAll14.Count > 0
                        ? LfStemShiShen(bhAll14[0].stem, dStem) : "";
                    string sSS14 = LfStemShiShen(lc14.stem, dStem);
                    bool isCur14 = lc14.startAge <= currentAge && lc14.endAge >= currentAge;
                    string marker14 = isCur14 ? " ★" : "";
                    sb.AppendLine($"| {lc14.startAge}-{lc14.endAge}{marker14} | {lc14.stem}{lc14.branch} | {sSS14} | {bSS14} | {lc14.score} | {lc14.level} |");
                }
                sb.AppendLine();
                sb.AppendLine($"（★ 現走大運，當前年齡 {currentAge} 歲，用神：{yongShenElem}，大忌：{jiShenElem}）");
                sb.AppendLine();

                var chartStemSS14 = new[] { yStemSS, mStemSS, "日主", hStemSS };

                // 二、逐運詳細論斷（天干期 + 地支期 + 紫微交叉）
                sb.AppendLine("二、逐運詳細論斷");
                sb.AppendLine();
                foreach (var lc in scored14)
                {
                    string stemElem14  = KbStemToElement(lc.stem);
                    bool hasHidden14 = LfBranchHiddenRatio.TryGetValue(lc.branch, out var bhLc14) && bhLc14.Count > 0;
                    string branchMainElem14 = hasHidden14 ? KbStemToElement(bhLc14[0].stem) : "";
                    string stemSS  = LfStemShiShen(lc.stem, dStem);
                    string branchSS = hasHidden14 ? LfStemShiShen(bhLc14[0].stem, dStem) : "";
                    bool isCurrent = lc.startAge <= currentAge && lc.endAge >= currentAge;
                    string curMark = isCurrent ? "【現走】" : "";

                    // 通根：地支藏干含天干五行（干支同氣，互相加持）
                    bool hasStemRoot14 = hasHidden14 && bhLc14.Any(h => KbStemToElement(h.stem) == stemElem14);
                    // 干克支 / 支克干
                    bool stemOverridesBranch14 = !string.IsNullOrEmpty(stemElem14) && !string.IsNullOrEmpty(branchMainElem14)
                        && LfElemOvercome.GetValueOrDefault(stemElem14, "") == branchMainElem14;
                    bool branchOverridesStem14 = !string.IsNullOrEmpty(branchMainElem14) && !string.IsNullOrEmpty(stemElem14)
                        && LfElemOvercome.GetValueOrDefault(branchMainElem14, "") == stemElem14;

                    bool stemIsGood14   = !string.IsNullOrEmpty(stemElem14)       && goodElems14.Contains(stemElem14);
                    bool stemIsBad14    = !string.IsNullOrEmpty(stemElem14)       && badElems14.Contains(stemElem14);
                    bool branchIsGood14 = !string.IsNullOrEmpty(branchMainElem14) && goodElems14.Contains(branchMainElem14);
                    bool branchIsBad14  = !string.IsNullOrEmpty(branchMainElem14) && badElems14.Contains(branchMainElem14);

                    // 天干期評分（天干為主×20，地支為輔×10）
                    // 通根 → 天干效力×1.5；支克干 → 天干效力×0.5；干克支 → 地支輔助×0.5
                    double stemMult14   = branchOverridesStem14  ? 0.5 : 1.0;
                    double branchMult14 = stemOverridesBranch14  ? 0.5 : 1.0;
                    double rootMult14   = hasStemRoot14 ? 1.5 : 1.0;
                    double stemDelta14 = stemIsGood14  ? 20 * stemMult14 * rootMult14
                                      : stemIsBad14   ? -20 * stemMult14 * rootMult14 : 0;
                    double branchAux14 = branchIsGood14 ? 10 * branchMult14
                                      : branchIsBad14  ? -10 * branchMult14 : 0;
                    int stemPeriodScore14 = (int)Math.Round(Math.Clamp(50 + stemDelta14 + branchAux14, 0, 100));

                    // 地支期評分（地支為主×20，天干為輔×10）
                    // 通根 → 地支效力×1.5（同一通根條件，干支同元互助）
                    double branchDelta14 = branchIsGood14 ? 20 * branchMult14 * rootMult14
                                        : branchIsBad14  ? -20 * branchMult14 * rootMult14 : 0;
                    double stemAux14 = stemIsGood14  ? 10 * stemMult14
                                    : stemIsBad14   ? -10 * stemMult14 : 0;
                    int branchPeriodScore14 = (int)Math.Round(Math.Clamp(50 + branchDelta14 + stemAux14, 0, 100));

                    // 紫微大限分析（以大運天干為主要流星，取大限中段年齡作年齡門控）
                    int midAge14 = (lc.startAge + lc.endAge) / 2;
                    int ziweiDecadeScore14 = hasZiwei ? DyCalcZiweiScore(lc.stem, palacesYdz, "", midAge14) : 50;

                    string stemCrossClass14   = DyCrossClass(stemPeriodScore14,   ziweiDecadeScore14);
                    string branchCrossClass14 = DyCrossClass(branchPeriodScore14, ziweiDecadeScore14);

                    // 大運標頭
                    sb.AppendLine($"{curMark}{lc.startAge}-{lc.endAge} 歲  大運：{lc.stem}{lc.branch}（天干{stemSS}·地支{branchSS}）  整體評分：{lc.score} 分·{lc.level}");

                    // 干支互動說明
                    var interParts14 = new List<string>();
                    if (hasStemRoot14)
                        interParts14.Add($"天干通根地支，{stemElem14}同氣相助，喜忌效力加倍");
                    if (stemOverridesBranch14)
                        interParts14.Add($"干（{stemElem14}）克支（{branchMainElem14}），地支力道減半");
                    else if (branchOverridesStem14)
                        interParts14.Add($"支（{branchMainElem14}）克干（{stemElem14}），天干力道減半");
                    if (interParts14.Count > 0)
                        sb.AppendLine($"干支互動：{string.Join("；", interParts14)}");

                    // 紫微大限命宮分析
                    if (hasZiwei && palacesYdz.ValueKind == JsonValueKind.Array)
                    {
                        var zDecades14 = DyGetOverlappingDecadePalaces(palacesYdz, lc.startAge, lc.endAge);
                        foreach (var (zDecPal14, zDecStem14, zDs14, zDe14) in zDecades14)
                        {
                            if (string.IsNullOrEmpty(zDecPal14)) continue;
                            string zPalStars14 = KbGetPalaceStars(palacesYdz, zDecPal14);
                            var (zGood14, zBad14) = DyGetPalaceAuxiliaryStars(palacesYdz, zDecPal14);
                            string zLu14   = string.IsNullOrEmpty(zDecStem14) ? "" : KbGetSiHuaPalace(zDecStem14, "化祿", palacesYdz);
                            string zQuan14 = string.IsNullOrEmpty(zDecStem14) ? "" : KbGetSiHuaPalace(zDecStem14, "化權", palacesYdz);
                            string zJi14   = string.IsNullOrEmpty(zDecStem14) ? "" : KbGetSiHuaPalace(zDecStem14, "化忌", palacesYdz);
                            // 輔星說明
                            var zAux14 = new List<string>();
                            if (zGood14.Count > 0) zAux14.Add($"六吉：{string.Join("、", zGood14)}");
                            if (zBad14.Count > 0)  zAux14.Add($"四煞：{string.Join("、", zBad14)}");
                            // 宮干四化
                            var zSiHua14 = new List<string>();
                            if (!string.IsNullOrEmpty(zLu14))   zSiHua14.Add($"化祿飛{zLu14}");
                            if (!string.IsNullOrEmpty(zQuan14)) zSiHua14.Add($"化權飛{zQuan14}");
                            if (!string.IsNullOrEmpty(zJi14))   zSiHua14.Add($"化忌飛{zJi14}");
                            // 主星質性
                            var goodMajor14   = new HashSet<string> { "紫微","天府","天相","天同","天梁","太陽","太陰" };
                            var activeMajor14 = new HashSet<string> { "七殺","破軍" };
                            var mixMajor14    = new HashSet<string> { "天機","武曲","廉貞","貪狼","巨門" };
                            var sArr14 = string.IsNullOrEmpty(zPalStars14) ? Array.Empty<string>()
                                : zPalStars14.Split(new[] { '、', ',', '，', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                            bool zHasGood14   = sArr14.Any(s => goodMajor14.Contains(s));
                            bool zHasActive14 = sArr14.Any(s => activeMajor14.Contains(s));
                            bool zHasMix14    = sArr14.Any(s => mixMajor14.Contains(s));
                            string zQuality14 =
                                string.IsNullOrEmpty(zPalStars14) ? "大限命宮空宮，借對宮主星論格局，大限方向較難聚焦。"
                              : zHasGood14 && zBad14.Count == 0 ? $"大限命宮{zPalStars14}入守，主星吉和穩健，大限格局有支撐。"
                              : zHasGood14 ? $"大限命宮{zPalStars14}入守，吉星有力；惟四煞同臨，格局有起伏，宜防意外突發。"
                              : zHasActive14 ? $"大限命宮{zPalStars14}入守，主星剛烈主變動，大限衝勁強，適合主動出擊，亦需防破耗。"
                              : zHasMix14 && zBad14.Count > 0 ? $"大限命宮{zPalStars14}入守，主星吉凶兩用，四煞同臨，大限阻力較多，宜謹慎行事。"
                              : $"大限命宮{zPalStars14}入守，主星吉凶需視宮干四化定論。";
                            var keyPals14 = new HashSet<string> { "命宮","財帛宮","官祿宮","遷移宮" };
                            if (!string.IsNullOrEmpty(zJi14) && keyPals14.Contains(zJi14))
                                zQuality14 += $" 宮干化忌入{zJi14.Replace("宮", "")}，大限此宮事項宜謹慎。";
                            else if (!string.IsNullOrEmpty(zLu14) && keyPals14.Contains(zLu14))
                                zQuality14 += $" 宮干化祿入{zLu14.Replace("宮", "")}，大限此宮事項有進益。";
                            // 輸出
                            string zDecRange14 = (zDs14 != lc.startAge || zDe14 != lc.endAge) ? $"（紫微大限 {zDs14}-{zDe14} 歲）" : "";
                            sb.AppendLine($"紫微大限{zDecRange14}：{zDecPal14}為大限命宮，主星：{(string.IsNullOrEmpty(zPalStars14) ? "空宮" : zPalStars14)}" +
                                (zAux14.Count > 0 ? $"，{string.Join("，", zAux14)}" : "") +
                                (!string.IsNullOrEmpty(zDecStem14) && zSiHua14.Count > 0 ? $"，宮干{zDecStem14}：{string.Join("、", zSiHua14)}" : ""));
                            sb.AppendLine($"  {zQuality14}");
                        }
                    }

                    // 天干期（前5年）
                    int stPeriodStart = lc.startAge;
                    int stPeriodEnd   = lc.startAge + 4;
                    string stCurMark = (currentAge >= stPeriodStart && currentAge <= stPeriodEnd) ? "★" : "";
                    sb.AppendLine($"  天干期（{stPeriodStart}-{stPeriodEnd} 歲）{stCurMark}：{lc.stem}（{stemSS}） 評分 {stemPeriodScore14} · 紫微 {ziweiDecadeScore14} · 綜合：{stemCrossClass14}");
                    string stemPeriodDesc14 =
                        stemPeriodScore14 >= 75 ? $"天干{lc.stem}（{stemSS}）喜用旺盛{(hasStemRoot14 ? "，更得地支通根加持" : "")}，此五年宜積極進取，把握機遇，拓展版圖的黃金窗口。"
                      : stemPeriodScore14 >= 55 ? $"天干{lc.stem}（{stemSS}）喜用有力{(branchOverridesStem14 ? "，雖受地支壓制，仍具動能" : "")}，宜順勢而為，主動佈局，積極推進。"
                      : stemPeriodScore14 >= 40 ? $"天干{lc.stem}（{stemSS}）力道平穩{(stemOverridesBranch14 ? "，地支力道已弱" : "")}，宜穩健行事，平中求吉，蓄勢積累。"
                      : stemPeriodScore14 >= 25 ? $"天干{lc.stem}（{stemSS}）忌神有力{(branchOverridesStem14 ? "，且受地支壓制，自身阻滯明顯" : "")}，宜低調守成，避免冒進決策。"
                      :                           $"天干{lc.stem}（{stemSS}）忌神旺盛{(hasStemRoot14 && stemIsBad14 ? "，更得通根強化，忌上加忌" : "")}，此五年壓力沉重，宜靜守蓄積，嚴防重大損失。";
                    sb.AppendLine($"  {stemPeriodDesc14}");
                    string stemCrossDesc14 = DyCrossDesc(stemCrossClass14, stemSS, branchSS, stemPeriodScore14, ziweiDecadeScore14);
                    sb.AppendLine($"  {stemCrossDesc14}");
                    sb.AppendLine();

                    // 地支期（後5年）
                    int brPeriodStart = lc.startAge + 5;
                    int brPeriodEnd   = lc.endAge;
                    string brCurMark = (currentAge >= brPeriodStart && currentAge <= brPeriodEnd) ? "★" : "";
                    sb.AppendLine($"  地支期（{brPeriodStart}-{brPeriodEnd} 歲）{brCurMark}：{lc.branch}（{branchSS}） 評分 {branchPeriodScore14} · 紫微 {ziweiDecadeScore14} · 綜合：{branchCrossClass14}");
                    string branchPeriodDesc14 =
                        branchPeriodScore14 >= 75 ? $"地支{lc.branch}（{branchSS}）喜用旺盛{(hasStemRoot14 ? "，干支同氣共振，力量更強" : "")}，此後五年宜積極深化，鞏固前期成果。"
                      : branchPeriodScore14 >= 55 ? $"地支{lc.branch}（{branchSS}）喜用有力{(stemOverridesBranch14 ? "，雖受天干壓制，根基仍穩" : "")}，宜鞏固成果，穩健推進。"
                      : branchPeriodScore14 >= 40 ? $"地支{lc.branch}（{branchSS}）力道平穩，宜守成蓄積，靜待時機，量力而為。"
                      : branchPeriodScore14 >= 25 ? $"地支{lc.branch}（{branchSS}）忌神有力{(branchOverridesStem14 && branchIsBad14 ? "，更克天干，干支雙忌" : "")}，宜低調守成，謹慎決策，減少風險暴露。"
                      :                             $"地支{lc.branch}（{branchSS}）忌神旺盛{(hasStemRoot14 && branchIsBad14 ? "，得天干同氣助長，忌上加忌" : "")}，此後五年持續承壓，宜極度守成，嚴控風險。";
                    sb.AppendLine($"  {branchPeriodDesc14}");
                    string branchCrossDesc14 = DyCrossDesc(branchCrossClass14, stemSS, branchSS, branchPeriodScore14, ziweiDecadeScore14);
                    if (branchCrossDesc14 != stemCrossDesc14)
                        sb.AppendLine($"  {branchCrossDesc14}");
                    sb.AppendLine();

                    // 地支六親宮位事項
                    string branchRelStr = LfGetBranchRelationsText(lc.branch, branches, dStem, yBranch, mBranch, hBranch);
                    if (!string.IsNullOrEmpty(branchRelStr))
                    {
                        sb.AppendLine($"  地支六親：大運地支 {lc.branch}（{branchSS}）");
                        sb.AppendLine(branchRelStr);
                    }

                    // 五大模組論斷（天干/地支引動/三干三支/空亡）
                    string stepAnalysis = LfDyStepAnalysis(
                        lc.stem, lc.branch, lc.startAge, lc.endAge, currentAge,
                        dStem, pillarStems14, pillarBranches14, pillarBranchSS14,
                        yongShenElem, jiShenElem, dayEmpty14, skipHeader: true);
                    if (!string.IsNullOrEmpty(stepAnalysis))
                        sb.AppendLine(stepAnalysis);
                    sb.AppendLine();
                }
            }

            // === Ch.13 開運指南 ===
            sb.AppendLine("【第十三章：開運指南】");
            sb.AppendLine();
            // 八字開運
            var elemDirV2 = new Dictionary<string, string> { {"木","東方"},{"火","南方"},{"土","中央"},{"金","西方"},{"水","北方"} };
            var openColorV2 = new Dictionary<string, string> { {"木","綠色、青色"},{"火","紅色、橙色"},{"土","黃色、米色"},{"金","白色、金色"},{"水","黑色、深藍色"} };
            var openMatV2 = new Dictionary<string, string> { {"木","植物盆栽、木製品、花草"},{"火","燈具、電器、紅色裝飾"},{"土","陶瓷、石材、黃色物品"},{"金","金屬飾品、白色物品"},{"水","水族箱、流水擺件、深色系"} };
            sb.AppendLine("【八字開運方向（用神五行）】");
            sb.AppendLine($"主喜用：{yongShenElem} → 方位：{elemDirV2.GetValueOrDefault(yongShenElem, "")}，色彩：{openColorV2.GetValueOrDefault(yongShenElem, "")}");
            sb.AppendLine($"輔助喜：{fuYiElem} → 方位：{elemDirV2.GetValueOrDefault(fuYiElem, "")}，材質：{openMatV2.GetValueOrDefault(yongShenElem, "")}");
            sb.AppendLine($"大忌：{jiShenElem} → 避免{elemDirV2.GetValueOrDefault(jiShenElem, "")}方向，避免{openColorV2.GetValueOrDefault(jiShenElem, "")}色系");
            sb.AppendLine();
            // 天乙貴人
            var tianYiMap = new Dictionary<string, string>
            {
                {"甲","丑未"},{"戊","丑未"},{"庚","丑未"},
                {"乙","子申"},{"己","子申"},
                {"丙","亥酉"},{"丁","亥酉"},
                {"壬","卯巳"},{"癸","卯巳"},
                {"辛","午寅"}
            };
            sb.AppendLine("【天乙貴人方向】");
            string tianYiBranches = tianYiMap.GetValueOrDefault(dStem, "");
            sb.AppendLine($"{dStem} 日主，天乙貴人在：{tianYiBranches}（見此地支方位或行此地支大運，貴人助力最強）");
            sb.AppendLine();
            // 大運貴人期
            var topScored = scored.OrderByDescending(s => s.score).Take(2).ToList();
            if (topScored.Count > 0)
            {
                sb.AppendLine("【大運最佳貴人期】");
                foreach (var ts in topScored)
                    sb.AppendLine($"  {ts.startAge}-{ts.endAge} 歲 {ts.stem}{ts.branch}（{ts.score}分）：此期貴人助力強，宜積極開創");
                sb.AppendLine();
            }
            // 紫微命宮開運
            if (hasZiwei && !string.IsNullOrEmpty(mingGongStars))
            {
                sb.AppendLine("【紫微命宮開運指引】");
                sb.AppendLine($"命宮主星：{mingGongStars}");
                if (siHua.TryGetValue("命宮化祿", out var mLuKv) && !string.IsNullOrEmpty(mLuKv.txt))
                    sb.AppendLine($"命宮化祿飛{mLuKv.pal}：{mLuKv.txt}（此宮位代表財運貴人方向）");
                if (siHua.TryGetValue("官祿化祿", out var oLuKv) && !string.IsNullOrEmpty(oLuKv.txt))
                    sb.AppendLine($"官祿化祿飛{oLuKv.pal}：{oLuKv.txt}（此宮位代表事業貴人方向）");
                sb.AppendLine();
            }
            // 最佳時辰 KB
            if (!string.IsNullOrEmpty(kb?.SpecialHours))
            {
                sb.AppendLine("【最佳時辰（古傳斷語）】");
                sb.AppendLine(kb.SpecialHours);
                sb.AppendLine();
            }
            // 行業建議
            sb.AppendLine("【適合行業綜合建議】");
            sb.Append(KbSanmenJobByElem(dmElem, pattern, yongShenElem, cfHuangliang, cfYangZhiYin));
            if (!string.IsNullOrEmpty(mingGongStars))
                sb.AppendLine($"  紫微命宮{mingGongStars}：建議依命宮星性質選擇職業方向");
            sb.AppendLine();

            // === Ch.14 出生環境（八字方位風水）===
            sb.AppendLine("【第十四章：出生環境・先天地理風水】");
            sb.AppendLine();
            sb.AppendLine("八字方點陣圖揭示命主懷胎時，父母受孕之地周圍的先天地理環境。");
            sb.AppendLine("年柱代表北方、月柱代表東方、日柱代表南方、時柱代表西方。");
            sb.AppendLine();
            sb.AppendLine($"              日柱（{dStem}{dBranch}）");
            sb.AppendLine($"                   南");
            sb.AppendLine($"                   ↑");
            sb.AppendLine($"月柱（{mStem}{mBranch}）  東 ← ─┼─ → 西  時柱（{hStem}{hBranch}）");
            sb.AppendLine($"                   ↓");
            sb.AppendLine($"                   北");
            sb.AppendLine($"              年柱（{yStem}{yBranch}）");
            sb.AppendLine();
            var fwPillars = new[]
            {
                (yStem+yBranch, "北方", "年柱"),
                (mStem+mBranch, "東方", "月柱"),
                (dStem+dBranch, "南方", "日柱"),
                (hStem+hBranch, "西方", "時柱")
            };
            foreach (var (fp, fwDir, fwName) in fwPillars)
            {
                sb.AppendLine($"▍{fwDir}（{fwName}：{fp}）");
                sb.AppendLine($"  {LfFengShuiPillarDesc(fp)}");
                sb.AppendLine();
            }
            sb.AppendLine("以上先天地理風水，指受孕當時所在地周圍的環境。驗證時，請參照父母受孕之地實際地貌加以對照。");
            sb.AppendLine();

            // === Ch.15 人生警示・趨吉避凶 ===
            sb.AppendLine("【第十五章：人生警示・趨吉避凶】");
            sb.AppendLine();
            sb.AppendLine("▍ 小人防範");
            sb.AppendLine(LfXiaoRenAnalysis(yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch, jiShenElem, dmElem));
            sb.AppendLine();
            sb.AppendLine("▍ 官司文書風險");
            sb.AppendLine(LfGuanSiAnalysis(yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch, jiShenElem, dmElem, bodyPct));
            sb.AppendLine();
            sb.AppendLine("▍ 車關時機");
            sb.AppendLine(LfCheGuanAnalysis(yBranch, mBranch, dBranch, hBranch, jiShenElem, dmElem));
            sb.AppendLine();
            sb.AppendLine("▍ 海外發展");
            sb.AppendLine(LfHaiWaiAnalysis(yBranch, mBranch, dBranch, hBranch, yongShenElem, jiShenElem, dmElem, hasZiwei, palacesYdz));
            sb.AppendLine();

            // === Ch.16 一生命運總評 ===
            sb.AppendLine("【第十六章：一生命運總評】");
            sb.AppendLine();
            if (scored.Count > 0)
            {
                double avg0_30 = scored.Where(s => s.startAge <= 30).Select(s => (double)s.score).DefaultIfEmpty(0).Average();
                double avg31_50 = scored.Where(s => s.startAge > 30 && s.startAge <= 50).Select(s => (double)s.score).DefaultIfEmpty(0).Average();
                double avg51 = scored.Where(s => s.startAge > 50).Select(s => (double)s.score).DefaultIfEmpty(0).Average();
                string DescAvg(double avg) => avg >= 65 ? "整體吉運強旺，宜積極開創" : avg >= 45 ? "整體平穩，有起有伏，宜穩中求進" : "整體考驗較多，宜低調蓄積，靜待轉機";
                sb.AppendLine($"前運（0-30 歲）平均 {avg0_30:F0} 分：{DescAvg(avg0_30)}");
                sb.AppendLine($"中運（31-50 歲）平均 {avg31_50:F0} 分：{DescAvg(avg31_50)}");
                sb.AppendLine($"後運（51 歲後）平均 {avg51:F0} 分：{DescAvg(avg51)}");
                sb.AppendLine();
                var best = scored.OrderByDescending(s => s.score).First();
                var worst = scored.OrderBy(s => s.score).First();
                sb.AppendLine($"人生最佳期：{best.startAge}-{best.endAge} 歲 {best.stem}{best.branch}（{best.score} 分）");
                sb.AppendLine($"人生考驗期：{worst.startAge}-{worst.endAge} 歲 {worst.stem}{worst.branch}（{worst.score} 分）");
                sb.AppendLine();
                // 財富等級
                double caiPct = wuXing.GetValueOrDefault(LfElemOvercome.GetValueOrDefault(dmElem, ""), 0);
                double guanPct = wuXing.GetValueOrDefault(LfElemOvercomeBy.GetValueOrDefault(dmElem, ""), 0);
                string caiLevel = caiPct >= 20 ? "豐厚" : caiPct >= 10 ? "中等" : "平淡";
                string guanLevel = guanPct >= 20 ? "顯達" : guanPct >= 10 ? "中等" : "平凡";
                sb.AppendLine($"財富等級：{caiLevel}（財星占 {caiPct:F0}%）  功名等級：{guanLevel}（官殺占 {guanPct:F0}%）");
                sb.AppendLine();
                sb.AppendLine($"命主喜走{yongShenElem}方位（{elemDirV2.GetValueOrDefault(yongShenElem, "")}），方能得天時地利。");
                sb.AppendLine($"趨吉避凶：謹慎避免【{jiShenElem}】方向，尤其在中凶/大凶運期間。");
            }
            sb.AppendLine();

            // === Ch.17 居家風水開運佈置 ===
            sb.AppendLine("【第十七章：居家風水開運佈置】");
            sb.AppendLine();
            sb.Append(KbSanmenFengShui(
                yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch,
                dmElem, bodyPct, yongShenElem, jiShenElem, wuXing, scored));
            sb.AppendLine();

            // === 表尾 ===
            sb.AppendLine("-----------------------------------------------------------------");
            sb.AppendLine("命理大師：玉洞子 | 玉洞子傳家寶典 v2.0");
            // 個人化：將正文「命主」替換為「XXX 先生/小姐」（保留「命主星/宮/垣/盤/：」不換）
            var reportResult = sb.ToString();
            if (!string.IsNullOrEmpty(userName))
                reportResult = System.Text.RegularExpressions.Regex.Replace(
                    reportResult, "命主(?![星宮垣盤：])", personRef);
            return reportResult;
        }

        // === 紫微格局偵測（依命宮主星+地支+四化） ===
        private static List<string> LfDetectZiweiGeJu(
            string mingStars, string mingBranch,
            HashSet<string> chartStars,
            string siHuaLuPalace, string siHuaQuanPalace, string siHuaKePalace,
            JsonElement palaces)
        {
            var matched = new List<string>();
            if (string.IsNullOrEmpty(mingStars)) { matched.Add("命無正曜格"); return matched; }

            bool Has(string s) => mingStars.Contains(s);

            // 星曜坐命格局
            if (Has("紫微") && Has("天府") && (mingBranch == "寅" || mingBranch == "申"))
                matched.Add("紫府同宮格");
            if (Has("太陽") && Has("太陰") && (mingBranch == "丑" || mingBranch == "未"))
                matched.Add("日月照壁格");
            if (Has("巨門") && Has("天機") && (mingBranch == "卯" || mingBranch == "酉"))
                matched.Add("巨機同臨格");
            if (Has("天同") && Has("太陰") && mingBranch == "子")
                matched.Add("月生滄海格");
            if (Has("武曲") && Has("貪狼") && (mingBranch == "丑" || mingBranch == "未"))
                matched.Add("貪武同行格");
            if (Has("左輔") && Has("右弼"))
                matched.Add("左右同宮格");
            if (Has("太陽") && Has("巨門") && (mingBranch == "寅" || mingBranch == "申"))
                matched.Add("巨日同宮格");
            if (Has("太陽") && mingBranch == "卯") matched.Add("日照雷門格");
            if (Has("太陽") && mingBranch == "午") matched.Add("日麗中天格");
            if (Has("太陰") && mingBranch == "亥") matched.Add("月朗天門格");
            if (Has("破軍") && (mingBranch == "子" || mingBranch == "午")) matched.Add("英星入廟格");
            if (Has("七殺") && (mingBranch == "子" || mingBranch == "午" || mingBranch == "寅" || mingBranch == "申"))
                matched.Add("七殺朝鬥格");
            if (Has("巨門") && (mingBranch == "子" || mingBranch == "午")) matched.Add("石中隱玉格");

            // 府相朝垣格：財帛宮及官祿宮各有天府或天相
            string wltStarsGj = KbGetPalaceStars(palaces, "財帛");
            string offStarsGj = KbGetPalaceStars(palaces, "官祿");
            bool wltHasFuXiang = wltStarsGj.Contains("天府") || wltStarsGj.Contains("天相");
            bool offHasFuXiang = offStarsGj.Contains("天府") || offStarsGj.Contains("天相");
            if (wltHasFuXiang && offHasFuXiang) matched.Add("府相朝垣格");

            // 四化格局（三方四正 = 命宮/財帛宮/官祿宮/遷移宮）
            var sf4z = new HashSet<string> { "命宮", "財帛宮", "官祿宮", "遷移宮" };
            bool InSF(string p) => sf4z.Contains(p) || sf4z.Contains(p + "宮");
            if (InSF(siHuaLuPalace) && InSF(siHuaQuanPalace) && InSF(siHuaKePalace)
                && !string.IsNullOrEmpty(siHuaLuPalace))
                matched.Add("三奇嘉會格");

            return matched;
        }

        // === 八字方位風水：單柱地理環境描述 ===
        private static string LfFengShuiPillarDesc(string pillar)
        {
            // 風水標誌較明顯的干支組合，優先使用
            var notableMap = new Dictionary<string, string>
            {
                ["甲寅"] = "有樹林、山林（陽性大樹如楊樹、榆樹、梧桐等），或木材相關場所（書店、紡織廠、農具行）",
                ["乙卯"] = "有樹林或公園（陰性樹木如槐樹、柳樹、竹子等），或花圃、園藝、手工業相關場所",
                ["甲子"] = "有一棵粗大顯眼的陽性大樹（水邊生長，樹形高大古老），或水邊有樹的環境",
                ["乙亥"] = "有一棵粗大顯眼的陰性大樹（水邊生長，如古柳、古槐等），或水邊有樹的環境",
                ["壬子"] = "有大水（水塘、水庫、河流、湖海、瀑布等），或寬闊繁忙的大馬路、交通幹道",
                ["壬辰"] = "有大水（坑塘、水庫、河流、湖面），或繁忙的交通要道、高速公路",
                ["癸亥"] = "有小水或靜水（水井、護城河、水溝、水渠等），或下水道、小型水利設施",
                ["癸丑"] = "有小水或靜水（水井、水溝、小河），或地下水道、廁坑、豬圈等水性場所",
                ["戊戌"] = "有山嶺、土丘、城牆、高牆、堤壩等高聳土石構造，或廟宇、庫房、窯場、鍋爐房",
                ["己未"] = "地勢低凹（溝渠、河道、低窪地形），或地勢下坡、凹陷的街道",
                ["己巳"] = "地勢低凹（水溝、河流、河床低陷），或與周圍相比明顯凹陷的地形",
                ["庚申"] = "有山嶺、石牆、水泥或金屬建築，或石橋、鐵橋、瀝青大路、金屬塔架",
                ["辛酉"] = "有石碑、石磨、假山、石景，或金屬器具店、珠寶業、武術館、交警單位等金屬性場所",
            };

            if (notableMap.TryGetValue(pillar, out string? notableDesc))
                return notableDesc;

            // 一般五行取象（依天干五行判斷）
            string stem = pillar.Length >= 1 ? pillar[..1] : "";
            string elem = stem switch
            {
                "甲" or "乙" => "木",
                "丙" or "丁" => "火",
                "戊" or "己" => "土",
                "庚" or "辛" => "金",
                "壬" or "癸" => "水",
                _ => ""
            };

            return elem switch
            {
                "木" => $"方位五行屬木，有樹木、樹林，或書店、木材行、紡織廠、農業相關場所",
                "火" => $"方位五行屬火，有高溫設施（鍋爐、煙囪）、工廠，或高大顯眼的建築物、電力設施",
                "土" => $"方位五行屬土，有土山、土丘、農田，或土石建材相關場所、廟宇、倉庫",
                "金" => $"方位五行屬金，有石山、金屬建築、道路，或金屬加工廠、礦場等金性場所",
                "水" => $"方位五行屬水，有水源、河流、水道，或流動性強的大路、交通運輸相關場所",
                _ => $"（干支五行待進一步分析，方位有相應地物或場所）",
            };
        }

        // === 大運地支宮位關係文字 ===
        private static string LfGetBranchRelationsText(string luckyBranch, string[] chartBranches,
            string dStem, string yBranch, string mBranch, string hBranch)
        {
            // 六沖、六合、三刑
            var chong6 = new Dictionary<string, string>
                { {"子","午"},{"午","子"},{"丑","未"},{"未","丑"},{"寅","申"},{"申","寅"},{"卯","酉"},{"酉","卯"},{"辰","戌"},{"戌","辰"},{"巳","亥"},{"亥","巳"} };
            var he6 = new Dictionary<string, string>
                { {"子","丑"},{"丑","子"},{"寅","亥"},{"亥","寅"},{"卯","戌"},{"戌","卯"},{"辰","酉"},{"酉","辰"},{"巳","申"},{"申","巳"},{"午","未"},{"未","午"} };

            var lines = new List<string>();
            var palaceMap = new[] {
                ("父母宮", mBranch), ("配偶宮", chartBranches[2]), ("子女宮", chartBranches[3]),
                ("兄弟宮", hBranch), ("年柱宮", yBranch)
            };
            string daySS = LfBranchHiddenRatio.TryGetValue(luckyBranch, out var bh) && bh.Count > 0
                ? LfStemShiShen(bh[0].stem, dStem) : "";

            foreach (var (palName, palBranch) in palaceMap)
            {
                if (string.IsNullOrEmpty(palBranch)) continue;
                var rels = new List<string>();
                if (chong6.GetValueOrDefault(luckyBranch) == palBranch) rels.Add("六沖");
                if (he6.GetValueOrDefault(luckyBranch) == palBranch) rels.Add("六合");
                if (rels.Count > 0)
                    lines.Add($"  與{palName}（{LfStemShiShen(LfBranchHiddenRatio.TryGetValue(palBranch,out var bh2)&&bh2.Count>0?bh2[0].stem:"",dStem)}·{palBranch}）{string.Join("、",rels)}，留意相關六親緣分變動。");
            }
            return string.Join("\n", lines);
        }

        // === 過三關·事業格局分析 ===

        private static string KbSanmenCareer(
            string yStem, string yBranch, string mStem, string mBranch,
            string dStem, string dBranch, string hStem, string hBranch,
            string yStemSS, string mStemSS, string hStemSS,
            string yBranchSS, string mBranchSS, string dBranchSS, string hBranchSS,
            string dmElem, string pattern, double bodyPct,
            string yongShenElem, string jiShenElem, Dictionary<string, double> wuXing)
        {
            var sb = new StringBuilder();

            // 外面 = 年月柱（社會/公共資源）；家裡 = 日時柱（個人/私領域）
            var outerSS = new[] { yStemSS, yBranchSS, mStemSS, mBranchSS };
            var innerSS = new[] { dBranchSS, hStemSS, hBranchSS };
            var allSS   = new[] { yStemSS, yBranchSS, mStemSS, mBranchSS, dBranchSS, hStemSS, hBranchSS };

            bool hasGuan      = allSS.Any(ss => ss == "正官");
            bool hasSha       = allSS.Any(ss => ss == "七殺");
            bool hasYin       = allSS.Any(ss => ss == "正印" || ss == "偏印");
            bool hasShiShen   = allSS.Any(ss => ss == "食神");
            bool hasShangGuan = allSS.Any(ss => ss == "傷官");
            bool hasShiShang  = hasShiShen || hasShangGuan;
            bool hasPianCai   = allSS.Any(ss => ss == "偏財");
            bool hasZhengCai  = allSS.Any(ss => ss == "正財");

            bool outerGuan = outerSS.Any(ss => ss == "正官");
            bool outerSha  = outerSS.Any(ss => ss == "七殺");
            bool outerYin  = outerSS.Any(ss => ss == "正印" || ss == "偏印");
            bool innerCai  = innerSS.Any(ss => ss == "偏財" || ss == "正財");

            bool isDayStemYang = "甲丙戊庚壬".Contains(dStem);

            // 列出年月/日時柱資源
            sb.AppendLine("【外面（社會資源）】");
            sb.AppendLine($"  年柱 {yStem}{yBranch}：年干={yStemSS}，年支={yBranchSS}");
            sb.AppendLine($"  月柱 {mStem}{mBranch}：月干={mStemSS}，月支={mBranchSS}");
            sb.AppendLine("【家裡（個人資源）】");
            sb.AppendLine($"  日支 {dBranch}：{dBranchSS}");
            sb.AppendLine($"  時柱 {hStem}{hBranch}：時干={hStemSS}，時支={hBranchSS}");
            sb.AppendLine();

            // 皇糧格局判斷
            bool isHuangliang = false;
            string huangliangType = "";

            // 殺印同宮（最強皇糧）
            bool shaYinSamePillar =
                (IsGuanSha(yStemSS) && IsYin(yBranchSS)) || (IsYin(yStemSS) && IsGuanSha(yBranchSS)) ||
                (IsGuanSha(mStemSS) && IsYin(mBranchSS)) || (IsYin(mStemSS) && IsGuanSha(mBranchSS)) ||
                (IsGuanSha(hStemSS) && IsYin(hBranchSS)) || (IsYin(hStemSS) && IsGuanSha(hBranchSS));

            if (shaYinSamePillar && (hasGuan || hasSha) && hasYin)
            {
                isHuangliang = true;
                huangliangType = "殺印同宮（官印相生同柱）";
            }
            else if ((hasGuan || hasSha) && hasYin)
            {
                isHuangliang = true;
                huangliangType = "化官生印（官印雙全）";
            }
            else if (hasSha && hasShiShen)
            {
                isHuangliang = true;
                huangliangType = "食神制殺（以食制殺，競爭性公職）";
            }
            else if (hasYin && (pattern == "建祿格" || pattern == "月刃格"))
            {
                isHuangliang = true;
                huangliangType = "印祿相隨（印比扶身，穩定薪水公職）";
            }

            // 陽制陰特殊格局（公安/司法/執法）
            bool isYangZhiYin = isDayStemYang && isHuangliang;

            // 自營傾向判定
            bool isZiYing        = pattern is "偏財格" or "食神格" or "傷官格";
            bool isZiYingPartial = hasPianCai && hasShiShang && !isHuangliang;

            // 格局結論
            if (isHuangliang)
            {
                sb.AppendLine("【事業格局判定·結論：公家政府命格】");
                sb.AppendLine($"  命局形成【{huangliangType}】，屬公家政府命格。");
                sb.AppendLine("  適合公職、體制內受薪、國營企業，或掛靠大型機構的穩定白領路線。");
                if (outerGuan || outerSha)
                    sb.AppendLine("  官殺星現於外面（社會），代表社會賦予的職位與名銜，天生與官場有緣。");
                if (outerYin)
                    sb.AppendLine("  印星出現在外面（社會），代表外部學歷/資格/長輩護持，有利晉升。");
                if (isYangZhiYin)
                    sb.AppendLine($"  命主 {dStem} 日主為陽干，具備公安、司法、軍警、執法方面的潛力。");
                sb.AppendLine();
                sb.AppendLine("【自營 vs 打工】結論：宜受薪/體制內，不宜冒進自營。");
            }
            else if (isZiYing)
            {
                sb.AppendLine("【事業格局判定·結論：自營創業命格】");
                sb.AppendLine($"  格局為【{pattern}】，屬自營/創業命格，財星活絡，資源由個人主導。");
                sb.AppendLine("  天生具備生意眼光與人際資源，適合自行創業、業務、自由業。");
                if (innerCai)
                    sb.AppendLine("  財星落於個人（家裡），賺的錢歸自己掌控，經商或自營為宜。");
                sb.AppendLine();
                sb.AppendLine("【自營 vs 打工】結論：宜自營創業，打工難以發揮潛力。");
            }
            else if (isZiYingPartial)
            {
                sb.AppendLine("【事業格局判定·結論：自營潛力，需等大運】");
                sb.AppendLine("  命局財星與食傷同現，有自營潛力，但需等用神大運到來方能發力。");
                sb.AppendLine("  建議先累積資本（打工/受薪），用神大運期再轉型自營。");
                sb.AppendLine();
                sb.AppendLine("【自營 vs 打工】結論：先打工積累，大運到來再轉型。");
            }
            else
            {
                sb.AppendLine("【事業格局判定·結論：技術/專業受薪命格】");
                sb.AppendLine("  命局財官兼具但無明顯公家政府或自營結構，屬中階受薪或技術型從業人員命格。");
                sb.AppendLine("  安穩打工、累積專業技能，方能穩步晉升。");
                sb.AppendLine();
                sb.AppendLine("【自營 vs 打工】結論：技術/專業受薪為正路，合夥自營需謹慎。");
            }

            return sb.ToString();
        }

        private static bool IsGuanSha(string ss) => ss == "正官" || ss == "七殺";
        private static bool IsYin(string ss)     => ss == "正印" || ss == "偏印";

        // 計算事業格局 flags（供 Ch.10 事業特質節直接呼叫）
        private static (bool isHuangliang, bool isYangZhiYin) KbCalcCareerFlags(
            string yStemSS, string yBranchSS, string mStemSS, string mBranchSS,
            string dBranchSS, string hStemSS, string hBranchSS,
            string dStem, string pattern)
        {
            var allSS = new[] { yStemSS, yBranchSS, mStemSS, mBranchSS, dBranchSS, hStemSS, hBranchSS };
            bool hasGuan    = allSS.Any(IsGuanSha);
            bool hasSha     = allSS.Any(s => s == "七殺");
            bool hasYin     = allSS.Any(IsYin);
            bool hasShiShen = allSS.Any(s => s == "食神");

            bool shaYinSamePillar =
                (IsGuanSha(yStemSS) && IsYin(yBranchSS)) || (IsYin(yStemSS) && IsGuanSha(yBranchSS)) ||
                (IsGuanSha(mStemSS) && IsYin(mBranchSS)) || (IsYin(mStemSS) && IsGuanSha(mBranchSS)) ||
                (IsGuanSha(hStemSS) && IsYin(hBranchSS)) || (IsYin(hStemSS) && IsGuanSha(hBranchSS));

            bool isHuangliang =
                (shaYinSamePillar && (hasGuan || hasSha) && hasYin) ||
                ((hasGuan || hasSha) && hasYin) ||
                (hasSha && hasShiShen) ||
                (hasYin && (pattern == "建祿格" || pattern == "月刃格"));

            bool isDayStemYang = "甲丙戊庚壬".Contains(dStem);
            bool isYangZhiYin  = isDayStemYang && isHuangliang;
            return (isHuangliang, isYangZhiYin);
        }

        private static string KbSanmenJobByElem(
            string dmElem, string pattern, string yongShenElem,
            bool isHuangliang, bool isYangZhiYin)
        {
            var sb = new StringBuilder();

            string patternJob = pattern switch
            {
                "正官格" => "政府機關、行政管理、財務會計、法務合規",
                "七殺格" => "軍警執法、管理顧問、外科醫療、競技運動",
                "正印格" => "教育學術、文化出版、醫療護理、策略研究",
                "偏印格" => "宗教藝術、技術研發、特殊技能、諮詢顧問",
                "正財格" => "金融財務、財務管理、穩定商業、行政辦事",
                "偏財格" => "商貿業務、房產仲介、投資理財、貿易進出口",
                "食神格" => "餐飲美食、技藝創作、設計藝術、娛樂表演",
                "傷官格" => "科技研發、創意設計、律師顧問、才藝自由業",
                "建祿格" => "專業技術、工程製造、自主事業、中小企業主",
                "月刃格" => "競爭性行業、業務銷售、投資合夥、體育競技",
                "曲直格" => "文教藝術、環保農林、設計創作、人文服務",
                "炎上格" => "科技電子、娛樂傳媒、能源照明、演藝創業",
                "稼穡格" => "農業地產、仲介保險、建材工程、農產食品",
                "從革格" => "金融財務、機械製造、法律軍警、珠寶精密",
                "潤下格" => "IT資訊、流通貿易、金融運輸、命理咨詢",
                _        => "多元發展，依大運時機選擇方向"
            };
            sb.AppendLine($"  格局取象：{patternJob}");

            string yongJob = yongShenElem switch
            {
                "木" => "教育文化、出版傳媒、木工園藝、環保文創、農林種植",
                "火" => "科技電子、能源照明、廚藝餐飲、娛樂演藝、電力暖氣",
                "土" => "建築房產、土木工程、農業農業、仲介保險、倉儲物流",
                "金" => "金融證券、機械製造、法律司法、珠寶冶金、軍警執法",
                "水" => "IT資訊、數據分析、航運物流、命理諮詢、傳播媒體",
                _    => "綜合型行業"
            };
            sb.AppendLine($"  用神{yongShenElem}取象：{yongJob}");

            if (isYangZhiYin)
                sb.AppendLine("  公安、司法、軍警、仲裁等執法類職業，命主天生具有優勢。");

            if (isHuangliang)
                sb.AppendLine("  提醒：優先選擇有編制、有保障的公家政府職位，切忌輕易放棄穩定。");

            return sb.ToString();
        }

        // === 六親緣分·婚姻深度鑑定 ===

        private static string KbSanmenSixRelatives(
            string yStem, string yBranch, string mStem, string mBranch,
            string dStem, string dBranch, string hStem, string hBranch,
            string yStemSS, string mStemSS, string hStemSS,
            string yBranchSS, string mBranchSS, string dBranchSS, string hBranchSS,
            string dmElem, string pattern, double bodyPct,
            string yongShenElem, string jiShenElem, Dictionary<string, double> wuXing,
            int gender, int birthYear,
            List<(string stem, string branch, string liuShen, int startAge, int endAge, int score, string level)> scored)
        {
            var sb = new StringBuilder();

            // 各柱六親宮位
            // 外面（社會/原生家庭）
            sb.AppendLine("【六親宮位分布】");
            sb.AppendLine($"  祖緣宮（年柱 {yStem}{yBranch}）：{yStemSS}／{yBranchSS}");
            sb.AppendLine($"  父母宮（月柱 {mStem}{mBranch}）：{mStemSS}／{mBranchSS}");
            sb.AppendLine($"  夫妻宮（日支 {dBranch}）：{dBranchSS}");
            sb.AppendLine($"  子女宮（時柱 {hStem}{hBranch}）：{hStemSS}／{hBranchSS}");
            sb.AppendLine();

            // 配偶星定義
            string spouseElem  = gender == 1
                ? LfElemOvercome.GetValueOrDefault(dmElem, "")   // 男：財星為妻
                : LfElemOvercomeBy.GetValueOrDefault(dmElem, ""); // 女：官殺為夫
            string spouseStar  = gender == 1 ? "妻星（財）" : "夫星（官殺）";
            double spousePct   = wuXing.GetValueOrDefault(spouseElem, 0);

            // 父星 = 偏財（異性父）；母星 = 正印（同性母）
            string fatherElem  = LfElemOvercome.GetValueOrDefault(dmElem, ""); // 偏財=父
            string motherElem  = LfGenByElem.GetValueOrDefault(dmElem, "");    // 印=母
            double fatherPct   = wuXing.GetValueOrDefault(fatherElem, 0);
            double motherPct   = wuXing.GetValueOrDefault(motherElem, 0);

            // 兄弟星 = 比劫
            double siblingPct  = wuXing.GetValueOrDefault(dmElem, 0);

            // 子女星 = 食傷
            string childElem   = LfElemGen.GetValueOrDefault(dmElem, "");
            double childPct    = wuXing.GetValueOrDefault(childElem, 0);

            // === 父母緣 ===
            sb.AppendLine("【父母緣】");
            string momLevel = motherPct >= 20 ? "深厚" : motherPct >= 10 ? "一般" : "較薄";
            string dadLevel = fatherPct >= 20 ? "深厚" : fatherPct >= 10 ? "一般" : "較薄";
            sb.AppendLine($"  母緣：印星（{motherElem}）占命局 {motherPct:F0}%，母子緣分【{momLevel}】。");
            sb.AppendLine(motherPct >= 15
                ? "  印星有力，母親影響力大，早年得母蔭護持，學歷資格有依靠。"
                : motherPct >= 8
                ? "  印星適中，與母親緣分尚可，情感有聯繫但自主性較強。"
                : "  印星偏弱，與母親聚少離多，或早年需自立，母緣較淡。");
            sb.AppendLine($"  父緣：財星（{fatherElem}）占命局 {fatherPct:F0}%，父子緣分【{dadLevel}】。");
            sb.AppendLine(fatherPct >= 15
                ? "  財星有力，父親資源豐厚，有機會繼承家業或得父親經濟支援。"
                : fatherPct >= 8
                ? "  財星適中，與父親關係正常，各自打拼為主。"
                : "  財星偏弱，父緣較淡，或父親早離、緣分不深，需靠自力更生。");
            // 月柱為父母宮，印/財入月吉
            bool yinInMonth   = mStemSS is "正印" or "偏印" || mBranchSS is "正印" or "偏印";
            bool caiInMonth   = mStemSS is "正財" or "偏財" || mBranchSS is "正財" or "偏財";
            if (yinInMonth)  sb.AppendLine("  印星現於父母宮（月柱），與母親關係尤為緊密，學識靠母方傳承。");
            if (caiInMonth)  sb.AppendLine("  財星現於父母宮（月柱），父親具財富資源，家境相對寬裕。");
            sb.AppendLine();

            // === 兄弟姐妹緣 ===
            sb.AppendLine("【兄弟姐妹緣】");
            string sibLevel = siblingPct >= 25 ? "深厚，兄弟姐妹情義濃" : siblingPct >= 15 ? "一般，情義有但各自獨立" : "較淡，獨立自主，少依賴手足";
            sb.AppendLine($"  比劫（{dmElem}）占命局 {siblingPct:F0}%，手足緣分【{sibLevel}】。");
            bool biInYear  = yStemSS is "比肩" or "劫財" || yBranchSS is "比肩" or "劫財";
            bool biInMonth = mStemSS is "比肩" or "劫財" || mBranchSS is "比肩" or "劫財";
            if (biInYear || biInMonth)
                sb.AppendLine("  比劫現於外面（年月），手足對命主社會發展有直接影響，需注意合夥或財務往來的得失。");
            else
                sb.AppendLine("  比劫多在個人（日時），手足情感在，但事業財務宜各自獨立，勿輕易合夥。");
            if (siblingPct >= 30)
                sb.AppendLine("  比劫過旺，手足競爭意識強，需防財務上的耗損或兄弟間利益衝突。");
            sb.AppendLine();

            // === 婚姻深度論斷 ===
            sb.AppendLine("【婚姻深度論斷】");
            sb.AppendLine($"  配偶星：{spouseStar}，五行 {spouseElem}，占命局 {spousePct:F0}%。");
            sb.AppendLine($"  夫妻宮（日支 {dBranch}）：{dBranchSS}。");

            // 配偶星強弱論斷
            string spouseLevel = spousePct >= 20 ? "旺" : spousePct >= 10 ? "適中" : "弱";
            if (spousePct >= 20)
                sb.AppendLine("  配偶星旺，感情機緣多，異性緣佳，擇偶條件好，婚姻資源豐富。");
            else if (spousePct >= 8)
                sb.AppendLine("  配偶星適中，緣分自然到來，婚姻情況穩定，無過多干擾。");
            else
                sb.AppendLine("  配偶星偏弱，感情緣分需耐心等候，或需主動創造機緣，切勿急於一時。");

            // 配偶星位置（在外面/家裡）
            bool spouseInOuter = false;
            bool spouseInInner = false;
            var allSS = new[] { yStemSS, yBranchSS, mStemSS, mBranchSS, dBranchSS, hStemSS, hBranchSS };
            var outerSSArr = new[] { yStemSS, yBranchSS, mStemSS, mBranchSS };
            var innerSSArr = new[] { dBranchSS, hStemSS, hBranchSS };
            if (gender == 1)
            {
                spouseInOuter = outerSSArr.Any(ss => ss is "正財" or "偏財");
                spouseInInner = innerSSArr.Any(ss => ss is "正財" or "偏財");
            }
            else
            {
                spouseInOuter = outerSSArr.Any(ss => ss is "正官" or "七殺");
                spouseInInner = innerSSArr.Any(ss => ss is "正官" or "七殺");
            }

            if (spouseInOuter && !spouseInInner)
                sb.AppendLine("  配偶星落於外面（社會），另一半多為社會上認識，工作或公眾場合中有緣相遇。");
            else if (spouseInInner && !spouseInOuter)
                sb.AppendLine("  配偶星落於個人（日時），另一半多為生活圈認識，青梅竹馬或私下牽線。");
            else if (spouseInOuter && spouseInInner)
                sb.AppendLine("  配偶星在外面與個人皆有，感情機緣來自多方，桃花旺盛。");

            // 日支吉凶
            bool dayBranchIsGood = dBranchSS == yongShenElem || dBranchSS == "正印" || dBranchSS == "食神";
            bool dayBranchIsBad  = dBranchSS == jiShenElem || dBranchSS == "七殺" || dBranchSS == "傷官";
            if (dayBranchIsGood)
                sb.AppendLine($"  夫妻宮日支 {dBranch} 屬吉，婚後生活穩固，另一半有助於命主運勢提升。");
            else if (dayBranchIsBad)
                sb.AppendLine($"  夫妻宮日支 {dBranch} 屬忌，婚姻中需多溝通磨合，宜選擇五行互補的另一半。");
            sb.AppendLine();

            // === 婚期推算 ===
            sb.AppendLine("【異性緣期】");
            var marriageLucks = scored.Where(lc =>
            {
                if (lc.startAge < 21) return false;  // 21歲以下大運不納入婚期推算
                string lcStemSS   = LfStemShiShen(lc.stem, dStem);
                string lcBrMs     = LfBranchHiddenRatio.TryGetValue(lc.branch, out var bh) && bh.Count > 0 ? bh[0].stem : "";
                string lcBranchSS = !string.IsNullOrEmpty(lcBrMs) ? LfStemShiShen(lcBrMs, dStem) : "";
                bool hasSpouse = gender == 1
                    ? lcStemSS is "正財" or "偏財" || lcBranchSS is "正財" or "偏財"
                    : lcStemSS is "正官" or "七殺" || lcBranchSS is "正官" or "七殺";
                return hasSpouse && lc.score >= 50;
            }).ToList();

            if (marriageLucks.Count > 0)
            {
                sb.AppendLine($"  命局中，以下大運期間配偶星入運，為感情婚姻最佳時機窗：");
                foreach (var lc in marriageLucks)
                    sb.AppendLine($"  - {lc.startAge}-{lc.endAge} 歲（{lc.stem}{lc.branch} 大運，評分 {lc.score} 分），{(gender == 1 ? "財星" : "官星")}有力，適合建立穩定感情或論及婚嫁。");
            }
            else
            {
                // 找評分最高且含配偶星的大運
                var bestLuck = scored.Where(lc =>
                {
                    if (lc.startAge < 21) return false;  // 21歲以下大運不納入婚期推算
                    string lcStemSS = LfStemShiShen(lc.stem, dStem);
                    string lcBrMs   = LfBranchHiddenRatio.TryGetValue(lc.branch, out var bh) && bh.Count > 0 ? bh[0].stem : "";
                    string lcBrSS   = !string.IsNullOrEmpty(lcBrMs) ? LfStemShiShen(lcBrMs, dStem) : "";
                    return gender == 1
                        ? lcStemSS is "正財" or "偏財" || lcBrSS is "正財" or "偏財"
                        : lcStemSS is "正官" or "七殺" || lcBrSS is "正官" or "七殺";
                }).OrderByDescending(lc => lc.score).FirstOrDefault();

                if (bestLuck != default)
                    sb.AppendLine($"  {bestLuck.startAge}-{bestLuck.endAge} 歲（{bestLuck.stem}{bestLuck.branch} 大運）配偶星入運，雖整體運勢需留意，但感情緣分仍可把握。");
                else
                    sb.AppendLine($"  命局配偶星較隱，感情緣分需靠流年時機（{spouseElem}年）主動創造，切勿守株待兔。");
            }

            // === 子女緣 ===
            sb.AppendLine();
            sb.AppendLine("【子女緣】");
            string childLevel = childPct >= 20 ? "深厚" : childPct >= 10 ? "一般" : "較薄";
            sb.AppendLine($"  子女星（食傷，{childElem}）占命局 {childPct:F0}%，子女緣分【{childLevel}】。");
            bool childInHour  = hStemSS is "食神" or "傷官" || hBranchSS is "食神" or "傷官";
            if (childPct >= 15)
                sb.AppendLine("  食傷有力，子女緣深，子女能力強，晚年受子女照顧。");
            else if (childPct >= 8)
                sb.AppendLine("  食傷適中，子女緣尚可，與子女情感平穩。");
            else
                sb.AppendLine("  食傷偏弱，子女緣較淡，或子女較少，亦可能晚婚晚育。");
            if (childInHour)
                sb.AppendLine($"  子女星現於時柱（子女宮），子女對命主晚年影響大，晚運多靠子女帶動。");
            sb.AppendLine();

            return sb.ToString();
        }

        // === 疾厄壽元鑑定 ===

        private static string KbSanmenHealthLongevity(
            string yStem, string mStem, string hStem,
            string yBranch, string mBranch, string dBranch, string hBranch,
            string dStem, string dmElem, double bodyPct,
            string yongShenElem, string jiShenElem,
            Dictionary<string, double> wuXing, string season, string seaLabel,
            List<(string stem, string branch, string liuShen, int startAge, int endAge, int score, string level)> scored,
            int currentAge = 0)
        {
            var sb = new StringBuilder();

            // 制神 = 剋住壞神（忌神）的五行
            string zhiShenElem = LfElemOvercomeBy.GetValueOrDefault(jiShenElem, "");
            double yongPct = wuXing.GetValueOrDefault(yongShenElem, 0);
            double jiPct   = wuXing.GetValueOrDefault(jiShenElem, 0);
            double zhiPct  = wuXing.GetValueOrDefault(zhiShenElem, 0);

            // === 元神/扶神/制神/壞神 ===
            sb.AppendLine("【元神·扶神·制神·壞神 四象定位】");
            sb.AppendLine($"  元神（生命核心）：{dmElem}（日主 {dStem}），元神{KbBodyStrengthShort(bodyPct)}。");
            sb.AppendLine($"  扶神（護持元神）：{yongShenElem}，占命局 {yongPct:F0}%，{(yongPct >= 15 ? "扶神有力，元神受護" : "扶神偏弱，護力不足")}。");
            sb.AppendLine($"  壞神（傷害命局）：{jiShenElem}，占命局 {jiPct:F0}%，{(jiPct >= 20 ? "壞神旺，對健康影響大" : "壞神偏弱，影響有限")}。");
            if (!string.IsNullOrEmpty(zhiShenElem))
                sb.AppendLine($"  制神（制住壞神）：{zhiShenElem}，占命局 {zhiPct:F0}%，{(zhiPct >= 10 ? "制神有力，壞神受制" : "制神弱，壞神難被壓制")}。");
            sb.AppendLine();

            // === 壽元強弱 ===
            sb.AppendLine("【壽元強弱判定】");
            string longevityLevel = bodyPct switch
            {
                >= 70 => "上等（元神極旺，生命力充沛，體質佳）",
                >= 55 => "中上（元神旺，體質良好，保養得當則壽元無虞）",
                >= 40 => "中等（元神適中，需注意忌神大運期的健康耗損）",
                >= 25 => "中下（元神偏弱，體質較虛，需積極保健）",
                _     => "偏弱（元神不足，早年需調養，避免大運忌神期透支）"
            };
            sb.AppendLine($"  壽元等級：【{longevityLevel}】");
            if (bodyPct >= 55)
                sb.AppendLine("  元神充足，先天體質良好，即使遭逢忌神大運也有較強的恢復力。");
            else if (bodyPct >= 40)
                sb.AppendLine("  元神尚可，體質平穩，忌神期需特別注意作息與飲食，不可過度消耗。");
            else
                sb.AppendLine($"  元神偏弱，宜多補養{yongShenElem}方向的食物與環境，嚴格避免{jiShenElem}方向進一步消耗。");
            sb.AppendLine();

            // === 乾坤戰識別 ===
            sb.AppendLine("【乾坤戰識別（陰陽五行激烈對抗）】");
            var sortedElems = wuXing.OrderByDescending(kv => kv.Value).ToList();
            bool qiankunZhan = false;
            string qiankunDesc = "";
            for (int i = 0; i < sortedElems.Count - 1 && !qiankunZhan; i++)
            {
                for (int j = i + 1; j < sortedElems.Count && !qiankunZhan; j++)
                {
                    string e1 = sortedElems[i].Key, e2 = sortedElems[j].Key;
                    double p1 = sortedElems[i].Value, p2 = sortedElems[j].Value;
                    bool isClash = LfElemOvercome.GetValueOrDefault(e1, "") == e2
                                || LfElemOvercome.GetValueOrDefault(e2, "") == e1;
                    if (isClash && p1 >= 25 && p2 >= 20 && Math.Abs(p1 - p2) <= 20)
                    {
                        qiankunZhan = true;
                        qiankunDesc = $"{e1}（{p1:F0}%）與{e2}（{p2:F0}%）激烈對抗";
                    }
                }
            }
            // 天干沖（四天干中任兩個形成對沖）
            var stemClashPairs = new HashSet<string> { "甲庚","庚甲","乙辛","辛乙","丙壬","壬丙","丁癸","癸丁" };
            var stems4 = new[] { yStem, mStem, dStem, hStem };
            bool stemClash = false;
            for (int i = 0; i < stems4.Length && !stemClash; i++)
                for (int j = i + 1; j < stems4.Length && !stemClash; j++)
                    if (stemClashPairs.Contains(stems4[i] + stems4[j])) stemClash = true;
            // 地支六沖
            var branchClashPairs = new HashSet<string> { "子午","午子","丑未","未丑","寅申","申寅","卯酉","酉卯","辰戌","戌辰","巳亥","亥巳" };
            var branches4 = new[] { yBranch, mBranch, dBranch, hBranch };
            bool branchClash = false;
            string clashBranches = "";
            for (int i = 0; i < branches4.Length && !branchClash; i++)
                for (int j = i + 1; j < branches4.Length && !branchClash; j++)
                    if (branchClashPairs.Contains(branches4[i] + branches4[j]))
                    { branchClash = true; clashBranches = $"{branches4[i]}{branches4[j]}"; }

            if (qiankunZhan)
                sb.AppendLine($"  命局形成五行對抗：{qiankunDesc}，氣機起伏不穩，情緒波動易影響健康。");
            if (stemClash)
                sb.AppendLine("  天干形成激沖，命局動盪，情緒壓力大，易有因壓力引發的身心症狀。");
            if (branchClash)
                sb.AppendLine($"  地支 {clashBranches} 形成對沖，臟腑能量不穩，每逢沖剋流年症狀易加重。");
            if (!qiankunZhan && !stemClash && !branchClash)
                sb.AppendLine("  命局陰陽五行相對和諧，無激烈對抗，體質穩定，氣機平順。");
            sb.AppendLine();

            // === 五行臟腑深度分析 ===
            sb.AppendLine("【五行臟腑深度分析】");
            var organMap = new Dictionary<string, (string organs, string symptoms, string care)>
            {
                { "木", ("肝膽、眼睛、筋骨", "眼疾、肝炎、筋骨酸痛、情緒焦慮", "規律作息勿熬夜、多吃綠色蔬菜護肝") },
                { "火", ("心臟、血液、血壓", "心悸、失眠、高血壓、貧血", "保持情緒平穩、適度有氧運動護心") },
                { "土", ("脾胃、消化系統", "消化不良、胃潰瘍、血糖偏高", "飲食規律、避免生冷寒涼、多吃黃色食物") },
                { "金", ("肺、大腸、皮膚", "咳嗽、鼻敏感、皮膚病、便秘", "注意呼吸道保健、遠離空汙、少吃辛辣") },
                { "水", ("腎臟、膀胱、骨骼", "腎虛、腰酸、骨質疏鬆、耳鳴", "充足睡眠養腎、多喝水、保暖護腰") },
            };
            string weakestElem  = wuXing.MinBy(kv => kv.Value).Key;
            foreach (var (elem, (organs, symptoms, care)) in organMap)
            {
                double pct = wuXing.GetValueOrDefault(elem, 0);
                if (pct >= 35)
                    sb.AppendLine($"  {elem}旺（{pct:F0}%）→ {organs}：易有{symptoms.Split('、')[0]}等亢進症狀，需節制。");
                else if (pct <= 8)
                    sb.AppendLine($"  {elem}弱（{pct:F0}%）→ {organs}：易有{symptoms}，保健建議：{care}。");
            }
            if (wuXing.Values.All(v => v > 8 && v < 35))
                sb.AppendLine("  五行分布均衡，臟腑整體和諧，日常保健維持即可。");
            sb.AppendLine($"  一生最需保養重點：{weakestElem}（{wuXing.GetValueOrDefault(weakestElem, 0):F0}%）對應之{organMap.GetValueOrDefault(weakestElem).organs}系統。");
            sb.AppendLine();

            // 體性調候影響
            if (seaLabel == "寒凍")
            {
                sb.AppendLine("【體性寒凍提醒】");
                sb.AppendLine("  命局體性偏寒，血液循環偏弱，腎虛、關節退化為主要風險。");
                sb.AppendLine("  宜居溫暖環境，多吃溫補食物（薑、桂圓、羊肉），避免冷飲冰品。");
                sb.AppendLine();
            }
            else if (seaLabel == "炎熱")
            {
                sb.AppendLine("【體性炎熱提醒】");
                sb.AppendLine("  命局體性偏熱，心火旺盛，心血管、高血壓、失眠為主要風險。");
                sb.AppendLine("  宜居涼爽環境，多吃清熱食物（冬瓜、綠豆、蓮子），避免燥熱刺激物。");
                sb.AppendLine();
            }

            // === 大運健康風險期 ===
            // 計算前一運起始年齡（同 Ch.14 邏輯），只顯示 >= 前一運的大運
            int riskCurIdx   = currentAge > 0 ? scored.FindIndex(lc => lc.startAge <= currentAge && lc.endAge >= currentAge) : 0;
            if (riskCurIdx < 0) riskCurIdx = scored.FindIndex(lc => lc.startAge > currentAge);
            if (riskCurIdx < 0) riskCurIdx = scored.Count - 1;
            int riskStartAge = riskCurIdx > 0 ? scored[riskCurIdx - 1].startAge : 0;

            sb.AppendLine("【大運健康風險期】");
            var riskLucks = scored.Where(lc =>
            {
                if (lc.startAge < riskStartAge) return false;
                string lcStemElem   = KbStemToElement(lc.stem);
                string lcBrMs       = LfBranchHiddenRatio.TryGetValue(lc.branch, out var bh) && bh.Count > 0 ? bh[0].stem : "";
                string lcBranchElem = !string.IsNullOrEmpty(lcBrMs) ? KbStemToElement(lcBrMs) : "";
                bool hasJi = lcStemElem == jiShenElem || lcBranchElem == jiShenElem;
                return hasJi && lc.score < 60;
            }).ToList();

            if (riskLucks.Count > 0)
            {
                sb.AppendLine($"  以下大運忌神（{jiShenElem}）入運且運分偏低，為健康需特別留意的時期：");
                foreach (var lc in riskLucks)
                    sb.AppendLine($"  - {lc.startAge}-{lc.endAge} 歲（{lc.stem}{lc.branch} 大運，評分 {lc.score} 分）：建議加強保健、定期體檢、作息規律。");
            }
            else
            {
                sb.AppendLine($"  大運中忌神（{jiShenElem}）整體影響尚可，無特別突出的高風險期。");
                sb.AppendLine("  維持健康習慣，定期保健即可。");
            }
            sb.AppendLine();
            sb.AppendLine("（以上為命理保健提醒，非醫療診斷，如有不適請就醫）");

            return sb.ToString();
        }

        private static string KbBodyStrengthShort(double bodyPct) => bodyPct switch
        {
            >= 70 => "極旺",
            >= 55 => "旺",
            >= 40 => "適中",
            >= 25 => "偏弱",
            _     => "弱"
        };

        // === 居家風水鑑定 ===

        private static string KbSanmenFengShui(
            string yStem, string yBranch, string mStem, string mBranch,
            string dStem, string dBranch, string hStem, string hBranch,
            string dmElem, double bodyPct,
            string yongShenElem, string jiShenElem,
            Dictionary<string, double> wuXing,
            List<(string stem, string branch, string liuShen, int startAge, int endAge, int score, string level)> scored)
        {
            var sb = new StringBuilder();

            // 天干類象（住宅環境原型）
            var stemImage = new Dictionary<string, string>
            {
                { "甲", "東方高地、樹木茂盛、背山面向寬闊之地，適合依山傍水的環境" },
                { "乙", "東方平地、花草樹木旁、安靜舒適的住宅區" },
                { "丙", "南方、向陽採光充足、鄰近熱鬧商業區或廣場" },
                { "丁", "南方、室內燈光溫暖、有爐灶廚房，溫馨居家型" },
                { "戊", "中央高地、山腹、田園廣闊，穩重厚實的環境" },
                { "己", "中央平原、農地、安靜的鄉村或郊區環境" },
                { "庚", "西方、鄰近道路交通要道、金屬工業區附近，利於外出發展" },
                { "辛", "西方、鬧中取靜、精緻住宅、商業區中的高樓" },
                { "壬", "北方、鄰近河流湖泊，或背水的開闊地帶" },
                { "癸", "北方、地勢低窪、近水源，需注意防潮防濕" },
            };

            // 地支類象（周邊環境特徵）
            var branchImage = new Dictionary<string, string>
            {
                { "子", "北方水氣重，宜注意防潮，鄰近水源有利但需防漏水" },
                { "丑", "東北方、倉儲型環境，周邊安靜，適合儲蓄型居所" },
                { "寅", "東北方山地，高地背山，採光好，生命力旺盛的社區" },
                { "卯", "正東方、住宅有大窗或大門向東，採光明亮，鄰近公園綠地" },
                { "辰", "東南方、潮濕土地或水庫附近，注意濕氣，適合近水的田園" },
                { "巳", "東南方、採光好、電氣設備齊全，近市場或商業區" },
                { "午", "正南方、陽光充足、門窗向南，開闊明亮，夏季較熱" },
                { "未", "西南方、田園花圃旁，安靜優雅的文化住宅區" },
                { "申", "正西方、鄰近道路或交通要道，金屬建材，現代感強" },
                { "酉", "正西方、鬧中有靜，廚房設備完善，精緻格局" },
                { "戌", "西北方高地、乾燥通風，適合山上或丘陵地帶的住宅" },
                { "亥", "正北方、近河流或地下水源，需注意防水防潮" },
            };

            // 五行方位
            var elemDir = new Dictionary<string, (string dir, string env, string avoid)>
            {
                { "木", ("東方", "鄰近公園、綠地、樹木茂盛之處；住宅宜木質裝潢，客廳擺植物，以東方窗採光為主", "西方金屬感重的環境、少植物的空間") },
                { "火", ("南方", "採光充足、向陽明亮；客廳宜偏暖色調，增加燈光亮度，擺放紅色或橙色裝飾", "北方陰暗潮濕的環境、光線不足的住宅") },
                { "土", ("中央、西南、東北", "平穩寬闊的住宅，不宜太多玻璃或流動感；黃色或米色系裝潢，擺放石材或陶瓷擺件", "頻繁搬遷、流動性過大的居住環境") },
                { "金", ("西方、西北", "整潔現代感的住宅，金屬裝潢，客廳簡約利落；白色或銀灰色調，擺放金屬或石材飾品", "東方木質感過重、雜亂零散的環境") },
                { "水", ("北方", "北方採光或水景設計；藍色或黑色系點綴，可在北方位置擺放水族箱或流水擺件", "南方正對烈日、過度乾燥炎熱的環境") },
            };

            // === 先天住宅原型分析（天干地支類象還原）===
            sb.AppendLine("【先天住宅原型（四柱天干地支類象）】");
            sb.AppendLine($"  年柱 {yStem}{yBranch}（出身/祖宅環境）：");
            sb.AppendLine($"    天干 {yStem}：{stemImage.GetValueOrDefault(yStem, "")}");
            sb.AppendLine($"    地支 {yBranch}：{branchImage.GetValueOrDefault(yBranch, "")}");
            sb.AppendLine($"  月柱 {mStem}{mBranch}（成長/父母宅環境）：");
            sb.AppendLine($"    天干 {mStem}：{stemImage.GetValueOrDefault(mStem, "")}");
            sb.AppendLine($"    地支 {mBranch}：{branchImage.GetValueOrDefault(mBranch, "")}");
            sb.AppendLine($"  日柱 {dStem}{dBranch}（本命/配偶宅環境）：");
            sb.AppendLine($"    天干 {dStem}：{stemImage.GetValueOrDefault(dStem, "")}");
            sb.AppendLine($"    地支 {dBranch}：{branchImage.GetValueOrDefault(dBranch, "")}");
            sb.AppendLine($"  時柱 {hStem}{hBranch}（晚年/子女宅環境）：");
            sb.AppendLine($"    天干 {hStem}：{stemImage.GetValueOrDefault(hStem, "")}");
            sb.AppendLine($"    地支 {hBranch}：{branchImage.GetValueOrDefault(hBranch, "")}");
            sb.AppendLine();

            // === 用神吉方（最適合居住的方位與環境）===
            sb.AppendLine("【用神吉方（最適合居住的方位與環境）】");
            if (elemDir.TryGetValue(yongShenElem, out var yongDir))
            {
                sb.AppendLine($"  命主用神為【{yongShenElem}】，最適合居住的方位：{yongDir.dir}。");
                sb.AppendLine($"  理想住宅特徵：{yongDir.env}。");
                sb.AppendLine($"  住宅座向建議：大門或主臥窗戶朝向{yongDir.dir}，引入{yongShenElem}氣助運。");
            }
            sb.AppendLine();

            // === 忌神凶方（應避免的方位與環境）===
            sb.AppendLine("【忌神凶方（應避免的方位與環境）】");
            if (elemDir.TryGetValue(jiShenElem, out var jiDir))
            {
                sb.AppendLine($"  命主忌神為【{jiShenElem}】，需迴避的方位：{jiDir.dir}。");
                sb.AppendLine($"  應避免：{jiDir.avoid}。");
                sb.AppendLine($"  若不得已住在忌神方位，可在該方位加強制化：擺放{yongShenElem}五行對應的物品化解。");
            }
            sb.AppendLine();

            // === 家居開運佈置建議 ===
            sb.AppendLine("【家居開運佈置建議】");
            string openColor = yongShenElem switch
            {
                "木" => "綠色、青色",
                "火" => "紅色、橙色、紫色",
                "土" => "黃色、米色、咖啡色",
                "金" => "白色、銀色、金色",
                "水" => "藍色、黑色、深灰色",
                _    => "中性色調"
            };
            string openMat = yongShenElem switch
            {
                "木" => "木質家具、植物盆栽、竹製品",
                "火" => "暖色燈光、蠟燭、紅色裝飾品",
                "土" => "陶瓷擺件、石材裝飾、黃色織品",
                "金" => "金屬擺件、銀色框架、白色石材",
                "水" => "水族箱、流水擺件、藍色掛畫",
                _    => "中性自然材質"
            };
            sb.AppendLine($"  開運色彩：{openColor}（{yongShenElem}行對應）");
            sb.AppendLine($"  開運材質：{openMat}");
            sb.AppendLine($"  主臥建議：床頭朝向{(elemDir.TryGetValue(yongShenElem, out var d) ? d.dir : "")}，避免床頭朝向{(elemDir.TryGetValue(jiShenElem, out var jd) ? jd.dir : "")}。");

            // 根據日主強弱給出特別建議
            if (bodyPct <= 35)
            {
                sb.AppendLine($"  元神偏弱建議：住宅宜選小而溫馨的格局，光線充足但不宜過大過空曠，聚氣為要。");
                sb.AppendLine($"  避免住在通風過強、過高或過空曠的住宅，以免洩氣。");
            }
            else if (bodyPct >= 65)
            {
                sb.AppendLine($"  元神旺盛建議：住宅可選寬闊開揚的格局，有助洩旺氣、達到平衡。");
                sb.AppendLine($"  高樓或面向開闊視野的住宅，有助於旺氣流動發展。");
            }
            sb.AppendLine();

            // === 搬遷吉運期 ===
            sb.AppendLine("【搬遷吉運期（大運時機）】");
            var moveLucks = scored.Where(lc => lc.score >= 70).OrderByDescending(lc => lc.score).Take(3).ToList();
            if (moveLucks.Count > 0)
            {
                sb.AppendLine("  以下大運為運勢高峰期，適合搬遷新宅或裝潢改造：");
                foreach (var lc in moveLucks)
                    sb.AppendLine($"  - {lc.startAge}-{lc.endAge} 歲（{lc.stem}{lc.branch} 大運，評分 {lc.score} 分）：此期搬遷或置產，有助鎖住好運。");
            }
            else
            {
                var bestLuck = scored.OrderByDescending(lc => lc.score).FirstOrDefault();
                if (bestLuck != default)
                    sb.AppendLine($"  最佳搬遷時機：{bestLuck.startAge}-{bestLuck.endAge} 歲（{bestLuck.stem}{bestLuck.branch} 大運，評分 {bestLuck.score} 分）。");
            }
            var avoidMoveLucks = scored.Where(lc => lc.score < 45).ToList();
            if (avoidMoveLucks.Count > 0)
            {
                sb.AppendLine("  以下大運運勢偏弱，不建議大幅搬遷或動土裝修：");
                foreach (var lc in avoidMoveLucks)
                    sb.AppendLine($"  - {lc.startAge}-{lc.endAge} 歲（{lc.stem}{lc.branch} 大運，評分 {lc.score} 分）：宜守舊，避免大動土木。");
            }

            return sb.ToString();
        }

        // === Lf Text Helpers ===

        // 計算四柱神殺
        private static List<string> LfGetBaziShenSha(
            string yStem, string yBranch, string mStem, string mBranch,
            string dStem, string dBranch, string hStem, string hBranch)
        {
            var result = new List<string>();
            var stems   = new[] { yStem, mStem, dStem, hStem };
            var branches = new[] { yBranch, mBranch, dBranch, hBranch };

            // 羊刃 - 日干對應陽刃地支
            var yangRen = new Dictionary<string,string>
            { {"甲","卯"},{"丙","午"},{"戊","午"},{"庚","酉"},{"壬","子"},
              {"乙","辰"},{"丁","未"},{"己","未"},{"辛","戌"},{"癸","丑"} };
            if (yangRen.TryGetValue(dStem, out var yrBr) && branches.Contains(yrBr))
                result.Add($"羊刃：日干 {dStem} 見 {yrBr}，刃星入命，個性剛強果決，宜武職或技術專業。");

            // 桃花 - 年支或日支對應桃花地支
            var taoHuaMap = new Dictionary<string,string>
            { {"子","酉"},{"丑","午"},{"寅","卯"},{"卯","子"},{"辰","酉"},{"巳","午"},
              {"午","卯"},{"未","子"},{"申","酉"},{"酉","午"},{"戌","卯"},{"亥","子"} };
            foreach (var baseBr in new[] { yBranch, dBranch })
            {
                if (taoHuaMap.TryGetValue(baseBr, out var thBr) && branches.Contains(thBr) && thBr != baseBr)
                {
                    result.Add($"桃花：{baseBr}生見 {thBr}，桃花入命，人緣極佳，異性緣旺，感情豐富。");
                    break;
                }
            }

            // 驛馬 - 年支或日支對應驛馬地支
            var yiMaMap = new Dictionary<string,string>
            { {"申","寅"},{"子","寅"},{"辰","寅"},
              {"亥","巳"},{"卯","巳"},{"未","巳"},
              {"寅","申"},{"午","申"},{"戌","申"},
              {"巳","亥"},{"酉","亥"},{"丑","亥"} };
            foreach (var baseBr in new[] { yBranch, dBranch })
            {
                if (yiMaMap.TryGetValue(baseBr, out var ymBr) && branches.Contains(ymBr))
                {
                    result.Add($"驛馬：{baseBr}生見 {ymBr}，驛馬入命，奔波勞碌，利於外出闖蕩或從事流動性工作。");
                    break;
                }
            }

            // 華蓋 - 年支或日支對應華蓋地支
            var huaGaiMap = new Dictionary<string,string>
            { {"申","辰"},{"子","辰"},{"辰","辰"},
              {"亥","未"},{"卯","未"},{"未","未"},
              {"寅","戌"},{"午","戌"},{"戌","戌"},
              {"巳","丑"},{"酉","丑"},{"丑","丑"} };
            foreach (var baseBr in new[] { yBranch, dBranch })
            {
                if (huaGaiMap.TryGetValue(baseBr, out var hgBr) && branches.Contains(hgBr))
                {
                    result.Add($"華蓋：{baseBr}生見 {hgBr}，華蓋入命，聰慧有才藝，帶有孤高之氣，宗教藝術緣深。");
                    break;
                }
            }

            // 孤辰寡宿 - 年支對應
            var guChenMap = new Dictionary<string,string>
            { {"寅","巳"},{"卯","巳"},{"辰","巳"},
              {"巳","申"},{"午","申"},{"未","申"},
              {"申","亥"},{"酉","亥"},{"戌","亥"},
              {"亥","寅"},{"子","寅"},{"丑","寅"} };
            if (guChenMap.TryGetValue(yBranch, out var gcBr) && branches.Contains(gcBr))
                result.Add($"孤辰：年支 {yBranch} 見 {gcBr}，孤辰入命，個性獨立，早年易孤單，晚年宜修身養性。");

            // 文昌 - 日干對應文昌地支
            var wenChangMap = new Dictionary<string,string>
            { {"甲","巳"},{"乙","午"},{"丙","申"},{"丁","酉"},{"戊","申"},
              {"己","酉"},{"庚","亥"},{"辛","子"},{"壬","寅"},{"癸","卯"} };
            if (wenChangMap.TryGetValue(dStem, out var wcBr) && branches.Contains(wcBr))
                result.Add($"文昌：日干 {dStem} 見 {wcBr}，文昌入命，聰明好學，利文筆考試，適合文教學術行業。");

            // 天乙貴人 - 日干對應天乙貴人地支（兩個）
            var tianYiMap = new Dictionary<string, string[]>
            { {"甲",new[]{"丑","未"}},{"乙",new[]{"子","申"}},{"丙",new[]{"亥","酉"}},
              {"丁",new[]{"亥","酉"}},{"戊",new[]{"丑","未"}},{"己",new[]{"子","申"}},
              {"庚",new[]{"丑","未"}},{"辛",new[]{"寅","午"}},{"壬",new[]{"卯","巳"}},
              {"癸",new[]{"卯","巳"}} };
            if (tianYiMap.TryGetValue(dStem, out var tyBrs))
            {
                var matched = tyBrs.Where(b => branches.Contains(b)).ToList();
                if (matched.Count > 0)
                    result.Add($"天乙貴人：日干 {dStem} 見 {string.Join("、",matched)}，天乙貴人入命，貴人多助，逢凶化吉，一生有貴人提攜。");
            }

            // 月德貴人 - 月支對應月德天干
            var yueDeMMap = new Dictionary<string,string>
            { {"寅","丙"},{"午","丙"},{"戌","丙"},
              {"申","壬"},{"子","壬"},{"辰","壬"},
              {"亥","甲"},{"卯","甲"},{"未","甲"},
              {"巳","庚"},{"酉","庚"},{"丑","庚"} };
            if (yueDeMMap.TryGetValue(mBranch, out var ydStem) && stems.Contains(ydStem))
                result.Add($"月德貴人：月支 {mBranch} 配天干 {ydStem}，月德貴人入命，得母系庇蔭，女性貴人緣深，化解官訟小人。");

            if (result.Count == 0)
                result.Add("四柱未見特定神殺入命。");

            return result;
        }

        private static string LfElemStems(string elem) => elem switch
        { "木"=>"甲乙","火"=>"丙丁","土"=>"戊己","金"=>"庚辛","水"=>"壬癸",_=>"" };

        private static string LfElemBranches(string elem) => elem switch
        { "木"=>"寅卯","火"=>"巳午","土"=>"辰戌丑未","金"=>"申酉","水"=>"亥子",_=>"" };

    // ======================================================================
    // 大運/流年分析通用方法（LfRun*）與大運專屬方法（LfDy*）
    // 方法論：memory/daiyun-liunian-methodology.md
    // ======================================================================

    // 單條引動結果
    private record BranchImpact(
        string RunChar,        // 大運/流年地支（顯示用）
        string RelationType,   // 六合/六沖/六害/三刑/三刑成局/爭合/六沖解合/三合/三會
        string Scenario,       // 1v1 / 解合 / 爭合 / 加入戰局 / 2+1成局 / 跨域成局
        string TargetBranch,   // 被引動的四柱地支（逗號分隔）
        string TargetPillar,   // 年/月/日/時（逗號分隔）
        string TargetSS,       // 被引動地支十神（+分隔）
        string PalaceName,     // 宮位名稱
        string FormedElement,  // 三合/三會成局五行；其他空字串
        string FavorStatus,    // 喜神/忌神/混合/中性
        string ImpactLevel     // ★重大運★ / 重大運
    );

    // ─── A：地支六關係引動（四情境）─────────────────────────────────────
    private static List<BranchImpact> LfRunBranchRelations(
        string runBranch,
        string[] pillarBranches,   // [年支, 月支, 日支, 時支]
        string[] pillarBranchSS,   // [年支十神, 月支十神, 日支十神, 時支十神]
        string yongShenElem, string jiShenElem,
        string dayunBranch = "")   // 僅流年分析傳入，啟用情境四
    {
        var results = new List<BranchImpact>();
        var labels  = new[] { "年", "月", "日", "時" };
        var palaces = new[] { "祖先宮", "父母宮", "夫妻宮", "子女宮" };

        // 取地支主氣元素 → 判斷喜忌
        string BranchFavor(string b)
        {
            if (!LfBranchHiddenRatio.TryGetValue(b, out var h)) return "中性";
            string e = KbStemToElement(h[0].stem);
            return e == yongShenElem ? "喜神" : e == jiShenElem ? "忌神" : "中性";
        }

        // ── 情境一：1v1 基本引動 ────────────────────────────
        for (int pi = 0; pi < 4; pi++)
        {
            string b = pillarBranches[pi];
            string rel = "";

            if (LfHe.TryGetValue(runBranch, out var heInfo) && heInfo.partner == b)
                rel = "六合";
            else if (LfChong.Contains(runBranch + b))
                rel = "六沖";
            else if (LfHai.Contains(runBranch + b))
                rel = "六害";
            else if (LfXing.Any(g => g.Contains(runBranch) && g.Contains(b) && b != runBranch))
                rel = "三刑";
            else if (b == runBranch && new[] { "辰","午","酉","亥" }.Contains(b))
                rel = "自刑";

            if (rel != "")
                results.Add(new BranchImpact(runBranch, rel, "1v1",
                    b, labels[pi], pillarBranchSS[pi], palaces[pi],
                    "", BranchFavor(b), "重大運"));
        }

        // ── 情境二：原有關係 + 運介入（解合 / 爭合 / 擴刑） ──
        for (int pi = 0; pi < 4; pi++)
        for (int pj = pi + 1; pj < 4; pj++)
        {
            string bi = pillarBranches[pi], bj = pillarBranches[pj];
            // 原有六合
            if (LfHe.TryGetValue(bi, out var heBI) && heBI.partner == bj)
            {
                // 運沖 bi → 解合
                if (LfChong.Contains(runBranch + bi))
                    results.Add(new BranchImpact(runBranch, "六沖解合", "解合",
                        $"{bi},{bj}", $"{labels[pi]},{labels[pj]}",
                        $"{pillarBranchSS[pi]}+{pillarBranchSS[pj]}",
                        $"{palaces[pi]}與{palaces[pj]}", "", BranchFavor(bi), "重大運"));
                // 運合 bi → 爭合
                if (LfHe.TryGetValue(runBranch, out var heRun) && heRun.partner == bi)
                    results.Add(new BranchImpact(runBranch, "爭合", "爭合",
                        $"{bi},{bj}", $"{labels[pi]},{labels[pj]}",
                        $"{pillarBranchSS[pi]}+{pillarBranchSS[pj]}",
                        $"{palaces[pi]}與{palaces[pj]}", "", BranchFavor(bi), "重大運"));
            }
        }
        // 原有刑 + 運擴為三刑
        foreach (var xGrp in LfXing.Where(g => g.Length == 3))
        {
            var inBazi = xGrp.Where(b => pillarBranches.Contains(b)).Distinct().ToList();
            if (inBazi.Count == 2 && xGrp.Contains(runBranch) && !inBazi.Contains(runBranch))
            {
                var inv = inBazi.Select(b => {
                    int i = Array.IndexOf(pillarBranches, b);
                    return (labels[i], pillarBranchSS[i], palaces[i]);
                }).ToList();
                string fav = new HashSet<string>(inBazi.Select(BranchFavor)).Count == 1
                    ? BranchFavor(inBazi[0]) : "混合";
                results.Add(new BranchImpact(runBranch, "三刑成局", "加入戰局",
                    string.Join(",", inBazi),
                    string.Join(",", inv.Select(x => x.Item1)),
                    string.Join("+", inv.Select(x => x.Item2)),
                    string.Join("、", inv.Select(x => x.Item3)),
                    "", fav, "★重大運★"));
            }
        }

        // ── 情境三：三合/三會 2+1 成局 ──────────────────────
        var allGroups = LfSanHe.Select(g => (g.branches, g.elem, "三合"))
            .Concat(LfSanHui.Select(g => (g.branches, g.elem, "三會")));
        foreach (var (grp, elem, gType) in allGroups)
        {
            if (!grp.Contains(runBranch)) continue;
            var needed = grp.Where(b => b != runBranch).ToList();
            var inBazi = needed.Where(b => pillarBranches.Contains(b)).Distinct().ToList();
            if (inBazi.Count < 2) continue;
            var inv = inBazi.Take(2).Select(b => {
                int i = Array.IndexOf(pillarBranches, b);
                return (labels[i], pillarBranchSS[i], palaces[i]);
            }).ToList();
            string fav = elem == yongShenElem ? "喜神" : elem == jiShenElem ? "忌神" : "中性";
            results.Add(new BranchImpact(runBranch, gType, "2+1成局",
                string.Join(",", inBazi.Take(2)),
                string.Join(",", inv.Select(x => x.Item1)),
                string.Join("+", inv.Select(x => x.Item2)),
                string.Join("、", inv.Select(x => x.Item3)),
                elem, fav, "★重大運★"));
        }

        // ── 情境四：跨域三合三會（流年傳入 dayunBranch 才啟用）──
        if (!string.IsNullOrEmpty(dayunBranch))
        {
            foreach (var (grp, elem, gType) in allGroups)
            {
                if (!grp.Contains(runBranch) || !grp.Contains(dayunBranch)) continue;
                if (runBranch == dayunBranch) continue;
                var remaining = grp.Where(b => b != runBranch && b != dayunBranch).ToList();
                var inBazi = remaining.Where(b => pillarBranches.Contains(b)).ToList();
                if (inBazi.Count == 0) continue;
                int idx = Array.IndexOf(pillarBranches, inBazi[0]);
                string fav = elem == yongShenElem ? "喜神" : elem == jiShenElem ? "忌神" : "中性";
                results.Add(new BranchImpact(runBranch, gType, "跨域成局",
                    inBazi[0], labels[idx], pillarBranchSS[idx], palaces[idx],
                    elem, fav, "★重大運★"));
            }
        }

        return results;
    }

    // ─── 白話說明生成 ───────────────────────────────────────────────────
    private static string LfRunBranchWhiteTalk(BranchImpact imp)
    {
        string ss = imp.TargetSS.Split('+')[0];   // 取第一個十神作主軸
        string palace = imp.PalaceName.Split('、')[0];
        string fav = imp.FavorStatus;

        // 十神 → 人事物類象
        string ssLabel = ss switch {
            "比肩" => "競爭朋友", "劫財" => "損財合夥", "食神" => "才藝子女口才",
            "傷官" => "換業官非才華突破", "偏財" => "偏財父緣投資",
            "正財" => "穩定財婚配（男）", "七殺" => "壓力官非異性（女）",
            "正官" => "升遷名譽婚配（女）", "偏印" => "學習宗教孤獨",
            "正印" => "貴人文書長輩護持", _ => ss
        };

        return (imp.Scenario, imp.RelationType, fav) switch {
            // 三合三會成局
            ("2+1成局" or "跨域成局", _, "喜神") =>
                $"{imp.FormedElement}局{imp.RelationType}成，喜神大旺，{palace}所主{ssLabel}爆發大機遇，此段全力把握。",
            ("2+1成局" or "跨域成局", _, "忌神") =>
                $"{imp.FormedElement}局{imp.RelationType}成，忌神大旺，{palace}所主{ssLabel}禍事大發，此段謹慎守成。",
            ("2+1成局" or "跨域成局", _, _) =>
                $"{imp.FormedElement}局{imp.RelationType}成，{palace}所主{ssLabel}出現重大變化，須密切觀察。",
            // 三刑成局
            ("加入戰局", "三刑成局", "喜神") =>
                $"三刑成局，{palace}喜神受多方衝擊，{ssLabel}有機遇但磨擦爭議難免，需防耗損。",
            ("加入戰局", "三刑成局", "忌神") =>
                $"三刑成局，{palace}忌神交戰，{ssLabel}衝突激烈，傷損官非需嚴防。",
            ("加入戰局", "三刑成局", _) =>
                $"三刑成局，{palace}多方磨擦，{ssLabel}糾纏不清，宜低調應對。",
            // 解合/爭合
            ("解合", _, _) =>
                $"大運沖開原有合局，{palace}關係生變，{ssLabel}方面出現重大轉折，情感或合作格局重組。",
            ("爭合", _, _) =>
                $"大運與原合局爭合，{palace}兩方拉鋸，{ssLabel}方面猶豫不決，難以兼顧，需擇一而行。",
            // 1v1 基本引動
            (_, "六合", "喜神") => $"喜神被合引動，{palace}所主{ssLabel}有機遇降臨，逢合易得貴緣，把握時機。",
            (_, "六合", "忌神") => $"忌神被合引動，{palace}所主{ssLabel}困擾被觸發，需防此類糾纏牽絆。",
            (_, "六合", _) => $"{palace}所主{ssLabel}被合引動，有緣分或合作機遇出現。",
            (_, "六沖", "喜神") => $"喜神被沖，{palace}所主{ssLabel}受衝擊，計劃易生變動，需謹慎保守。",
            (_, "六沖", "忌神") => $"忌神被沖破，{palace}所主{ssLabel}舊有阻礙有望突破，把握改變機會。",
            (_, "六沖", _) => $"{palace}所主{ssLabel}受沖，此方面易生衝突變動。",
            (_, "六害", "喜神") => $"喜神受害，{palace}所主{ssLabel}暗損，防此方面無形耗損。",
            (_, "六害", "忌神") => $"忌神受害，{palace}所主{ssLabel}受損，有助削弱不利因素。",
            (_, "六害", _) => $"{palace}所主{ssLabel}受害，易有暗損阻礙。",
            (_, "三刑" or "自刑", "喜神") => $"喜神受刑，{palace}所主{ssLabel}磨擦糾紛，防傷損官非。",
            (_, "三刑" or "自刑", "忌神") => $"忌神受刑，{palace}所主{ssLabel}衝突更烈，此類問題特別突出。",
            (_, "三刑" or "自刑", _) => $"{palace}所主{ssLabel}受刑，此方面有磨擦爭議，宜低調處理。",
            (_, "六沖解合", _) => $"原合局被沖開，{palace}所主{ssLabel}關係斷裂，情感或合作出現轉折。",
            _ => $"{palace}所主{ssLabel}受{imp.RelationType}影響，此方面出現變化。"
        };
    }

    // ─── 格式化單條引動輸出（含空亡/神煞/白話）──────────────────────────
    private static string LfRunFormatImpact(
        BranchImpact imp,
        string[] dayEmpty,
        string[] pillarBranches,
        string runLabel,        // "運（申）" 或 "年（申）"
        string crossDayun = "")  // 跨域成局時顯示大運地支
    {
        var sb = new System.Text.StringBuilder();
        bool isBig = imp.ImpactLevel.StartsWith("★");
        string mark = isBig ? " ★重大運★" : "";
        string crossNote = imp.Scenario == "跨域成局" && !string.IsNullOrEmpty(crossDayun)
            ? $"＋運（{crossDayun}）" : "";

        sb.AppendLine($"  ┌─{mark} {runLabel} {imp.RelationType} 四柱{imp.TargetPillar}宮（{imp.TargetBranch}）{crossNote}");
        sb.AppendLine($"  │   十神：{imp.TargetSS}　宮位：{imp.PalaceName}");

        // 空亡判斷
        var kbBranches = imp.TargetBranch.Split(',').Where(b => dayEmpty.Contains(b)).ToList();
        sb.AppendLine(kbBranches.Count > 0
            ? $"  │   空亡：（{string.Join(",", kbBranches)}）落空亡，力量減半"
            : "  │   空亡：無");

        // 神煞判斷（以四柱各地支為SKYNO，查被引動地支的神煞）
        var starSet = new HashSet<string>();
        foreach (var tb in imp.TargetBranch.Split(','))
            foreach (var pb in pillarBranches)
                if (DiZhiShenShaMap.TryGetValue(pb, out var m) && m.TryGetValue(tb, out var ss))
                    foreach (var s in ss) starSet.Add(s);
        sb.AppendLine(starSet.Count > 0
            ? $"  │   神煞：{string.Join("、", starSet.Select(s => $"({s})"))}引動"
            : "  │   神煞：無");

        // 白話
        sb.AppendLine($"  │   白話：{LfRunBranchWhiteTalk(imp)}");
        sb.AppendLine("  └─");
        return sb.ToString();
    }

    // ─── D：三干三支成局檢查 ─────────────────────────────────────────────
    private static string LfRunSanGanZhiCheck(
        string runStem, string runBranch,
        string[] pillarStems, string[] pillarBranches)
    {
        var sb = new System.Text.StringBuilder();

        // 三干
        var stemCounts = pillarStems.GroupBy(s => s).ToDictionary(g => g.Key, g => g.Count());
        if (stemCounts.TryGetValue(runStem, out int sc) && sc >= 2)
        {
            string prefix = sc == 2 ? "三" : "四";
            string key = $"{prefix}{runStem}";
            if (LfSanGanMap.TryGetValue(key, out var desc))
            {
                sb.AppendLine($"  八字已有 {sc} 個「{runStem}」，運（{runStem}）形成第 {sc+1} 個");
                sb.AppendLine($"  → 古傳 {key}：{desc}");
            }
        }

        // 三支
        var branchCounts = pillarBranches.GroupBy(b => b).ToDictionary(g => g.Key, g => g.Count());
        if (branchCounts.TryGetValue(runBranch, out int bc) && bc >= 2)
        {
            string key = $"三{runBranch}";
            if (LfSanZhiMap.TryGetValue(key, out var desc))
            {
                sb.AppendLine($"  八字已有 {bc} 個「{runBranch}」，運（{runBranch}）形成第 {bc+1} 個");
                sb.AppendLine($"  → 古傳 {key}：{desc}");
            }
        }

        return sb.ToString().Trim();
    }

    // ─── E：大運干支關係論斷（同氣/干克支/支克干）────────────────────────
    private static string LfDyStemBranchRelation(
        string dyStem, string dyStemSS, string dyStemElem,
        string dyBranch, string dyBranchSS, string dyBranchElem,
        int startAge, int endAge)
    {
        bool sameQi = dyStemElem == dyBranchElem
            || LfElemGen.GetValueOrDefault(dyStemElem) == dyBranchElem
            || LfElemGen.GetValueOrDefault(dyBranchElem) == dyStemElem;
        bool stemKills  = LfElemOvercome.GetValueOrDefault(dyStemElem)  == dyBranchElem;
        bool branchKills = LfElemOvercome.GetValueOrDefault(dyBranchElem) == dyStemElem;

        if (sameQi)
            return $"  天干（{dyStem}·{dyStemElem}）與地支（{dyBranch}·{dyBranchElem}）五行同氣相輔\n" +
                   $"  → 整段 {startAge}-{endAge} 歲方向一致，{dyStemSS}與{dyBranchSS}同向發力，吉凶十年貫串。";
        if (stemKills)
            return $"  天干（{dyStem}·{dyStemElem}）克地支（{dyBranch}·{dyBranchElem}）\n" +
                   $"  → 前五年（{startAge}-{startAge+4}歲）天干{dyStemSS}事件較主導；" +
                   $"後五年（{startAge+5}-{endAge}歲）地支{dyBranchSS}影響漸強。";
        if (branchKills)
            return $"  地支（{dyBranch}·{dyBranchElem}）克天干（{dyStem}·{dyStemElem}）\n" +
                   $"  → 前五年（{startAge}-{startAge+4}歲）地支{dyBranchSS}先發；" +
                   $"後五年（{startAge+5}-{endAge}歲）天干{dyStemSS}反作用漸顯。";

        return $"  天干（{dyStem}·{dyStemElem}）與地支（{dyBranch}·{dyBranchElem}）各行其是\n" +
               $"  → 前五年天干{dyStemSS}、後五年地支{dyBranchSS}分段主導。";
    }

    // ─── 大運單步完整分析 ────────────────────────────────────────────────
    private static string LfDyStepAnalysis(
        string dyStem, string dyBranch, int startAge, int endAge, int currentAge,
        string dStem,   // 日主天干（用於計算大運天干十神）
        string[] pillarStems, string[] pillarBranches,
        string[] pillarBranchSS,
        string yongShenElem, string jiShenElem,
        string[] dayEmpty,
        bool skipHeader = false)
    {
        // 大運天干十神
        string dyStemSS  = LfStemShiShen(dyStem, dStem);
        string dyStemElem = KbStemToElement(dyStem);

        // 大運地支主氣元素與十神
        string dyBranchMainStem = LfBranchHiddenRatio.TryGetValue(dyBranch, out var h)
            ? h[0].stem : "";
        string dyBranchSS  = string.IsNullOrEmpty(dyBranchMainStem) ? ""
            : LfStemShiShen(dyBranchMainStem, dStem);
        string dyBranchElem = string.IsNullOrEmpty(dyBranchMainStem) ? ""
            : KbStemToElement(dyBranchMainStem);

        bool isCurrent = currentAge >= startAge && currentAge <= endAge;
        var sb = new System.Text.StringBuilder();

        if (!skipHeader)
        {
            string currentMark = isCurrent ? $"  ← 目前行運（{currentAge}歲）" : "";
            sb.AppendLine($"【{startAge}-{endAge}歲｜{dyStem}{dyBranch} 大運】{currentMark}");
            sb.AppendLine();
        }

        // ▌一、天干論斷
        sb.AppendLine("▌一、天干論斷");
        sb.AppendLine($"  運（{dyStem}）為{dyStemSS}，五行{dyStemElem}");
        sb.AppendLine(LfDyStemBranchRelation(
            dyStem, dyStemSS, dyStemElem,
            dyBranch, dyBranchSS, dyBranchElem,
            startAge, endAge));
        sb.AppendLine();

        // ▌二、地支引動逐條分析
        sb.AppendLine("▌二、地支引動逐條分析");
        // pillarStemSS 在此方法不傳入，以空串替代（只需要 branchSS）
        var impacts = LfRunBranchRelations(
            dyBranch, pillarBranches, pillarBranchSS,
            yongShenElem, jiShenElem);
        if (impacts.Count == 0)
            sb.AppendLine("  （大運地支與四柱無刑沖害合，此段運勢較為平穩）");
        else
            foreach (var imp in impacts)
                sb.Append(LfRunFormatImpact(imp, dayEmpty, pillarBranches, $"運（{dyBranch}）"));
        sb.AppendLine();

        // ▌三、三干三支
        sb.AppendLine("▌三、三干/三支成局");
        string sanCheck = LfRunSanGanZhiCheck(dyStem, dyBranch, pillarStems, pillarBranches);
        sb.AppendLine(string.IsNullOrEmpty(sanCheck) ? "  （無三干三支成局）" : sanCheck);
        sb.AppendLine();

        // ▌四、空亡
        sb.AppendLine("▌四、空亡");
        if (dayEmpty.Contains(dyBranch))
        {
            string fav = LfBranchHiddenRatio.TryGetValue(dyBranch, out var hd)
                ? KbStemToElement(hd[0].stem) == yongShenElem ? "喜神"
                : KbStemToElement(hd[0].stem) == jiShenElem   ? "忌神" : "中性"
                : "中性";
            sb.AppendLine($"  運（{dyBranch}）落入空亡（旬空：{dayEmpty[0]}{dayEmpty[1]}）");
            sb.AppendLine(fav == "喜神"
                ? "  → 喜神行空亡運，十年努力易名存實亡，謀事宜保守，成效打折。"
                : fav == "忌神"
                ? "  → 忌神行空亡運，凶性虛而不實，十年禍事力道減半。"
                : "  → 此步行空亡，整體謀事宜低調保守。");
        }
        else
        {
            sb.AppendLine("  （無空亡）");
        }
        sb.AppendLine();

        return sb.ToString().TrimEnd();
    }

    // ─── 流年單年完整分析 ──────────────────────────────────────────────────
    private static string LfLnYearAnalysis(
        string lyStem, string lyBranch, int lyYear, string zodiac,
        string dStem,
        string[] pillarStems, string[] pillarBranches,
        string[] pillarBranchSS,
        string yongShenElem, string jiShenElem,
        string[] dayEmpty,
        string dayunBranch = "",   // 當前大運地支，傳入啟用情境四跨域三合
        bool skipHeader = false)
    {
        string lyStemSS   = LfStemShiShen(lyStem, dStem);
        string lyBranchMainStem = LfBranchHiddenRatio.TryGetValue(lyBranch, out var h)
            ? h[0].stem : "";
        string lyBranchSS = string.IsNullOrEmpty(lyBranchMainStem) ? ""
            : LfStemShiShen(lyBranchMainStem, dStem);

        // 流年天干事件類型
        string stemEvent = lyStemSS switch {
            "比肩" => "競爭加劇、防破財、兄弟朋友關係變動",
            "劫財" => "破財損失、防合夥糾紛、朋友之患",
            "食神" => "才藝發揮、飲食豐足、子女喜事",
            "傷官" => "換工作、官非口舌、創業突破",
            "偏財" => "偏財機遇、父緣事件、投資出現機會",
            "正財" => "穩定財源、正職收入、婚配緣分（男）",
            "七殺" => "壓力考驗、小人官非、挑戰突發",
            "正官" => "升遷機會、名譽事件、婚配緣分（女）",
            "偏印" => "學習進修、宗教靜思、易孤獨多慮",
            "正印" => "貴人資助、文書資格、長輩護持",
            _ => lyStemSS
        };

        var sb = new System.Text.StringBuilder();

        if (!skipHeader)
        {
            sb.AppendLine($"【{lyYear} {lyStem}{lyBranch}（{zodiac}年）】");
            sb.AppendLine();
        }

        // ▌一、流年天干
        sb.AppendLine("▌一、流年天干");
        sb.AppendLine($"  年干（{lyStem}）為{lyStemSS}，當年主要事件類型：{stemEvent}");
        sb.AppendLine();

        // ▌二、地支引動逐條分析
        sb.AppendLine("▌二、地支引動逐條分析");
        var impacts = LfRunBranchRelations(
            lyBranch, pillarBranches, pillarBranchSS,
            yongShenElem, jiShenElem, dayunBranch);
        if (impacts.Count == 0)
            sb.AppendLine("  （流年地支與四柱無刑沖害合，本年整體平穩）");
        else
            foreach (var imp in impacts)
                sb.Append(LfRunFormatImpact(imp, dayEmpty, pillarBranches,
                    $"年（{lyBranch}）", dayunBranch));
        sb.AppendLine();

        // ▌三、三干/三支成局
        sb.AppendLine("▌三、三干/三支成局");
        string sanCheck = LfRunSanGanZhiCheck(lyStem, lyBranch, pillarStems, pillarBranches);
        sb.AppendLine(string.IsNullOrEmpty(sanCheck) ? "  （無三干三支成局）" : sanCheck);
        sb.AppendLine();

        // ▌四、空亡
        sb.AppendLine("▌四、空亡");
        if (dayEmpty.Contains(lyBranch))
        {
            string fav = LfBranchHiddenRatio.TryGetValue(lyBranch, out var hd)
                ? KbStemToElement(hd[0].stem) == yongShenElem ? "喜神"
                : KbStemToElement(hd[0].stem) == jiShenElem   ? "忌神" : "中性"
                : "中性";
            sb.AppendLine($"  年（{lyBranch}）落入空亡（旬空：{dayEmpty[0]}{dayEmpty[1]}）");
            sb.AppendLine(fav == "喜神"
                ? "  → 喜神行空亡年，謀事易名存實亡，成效打折，不宜大動。"
                : fav == "忌神"
                ? "  → 忌神行空亡年，凶性虛而不實，禍事力道減半。"
                : "  → 今年行空亡，謀事宜低調保守，勿輕易大動。");
        }
        else
        {
            sb.AppendLine("  （無空亡）");
        }
        sb.AppendLine();

        return sb.ToString().TrimEnd();
    }

    // 四柱神煞論斷（來源：DB 地支星剎 + 天干星剎，四 KIND 全部交叉）
    private static string LfShenShaStarDesc(string star) => star switch
    {
        "將星" => "領導掌權", "劫煞" => "破財小人", "災煞" => "傷病意外",
        "驛馬" => "奔波移動", "桃花" => "人緣感情", "亡神" => "損耗消散",
        "華蓋" => "孤高藝術", "寡宿" => "孤寂清冷", "孤辰" => "孤立不群",
        "紅鸞" => "喜慶婚嫁", "天喜" => "吉慶喜事", "乙貴" => "天乙貴人",
        "文昌" => "文才學術", "干祿" => "本命食祿", "羊刃" => "剛烈衝動",
        "飛刃" => "刑剋傷損", "血刃" => "血光意外", "紅艷" => "異性桃花",
        "學士" => "才學聰穎", "路空" => "虛而不實", _ => ""
    };

    private static string LfShenSha(
        string yStem, string yBranch, string mStem, string mBranch,
        string dStem, string dBranch, string hStem, string hBranch,
        string yStemSS, string mStemSS, string hStemSS,
        string yBranchSS, string mBranchSS, string dBranchSS, string hBranchSS)
    {
        var stems   = new[] { yStem,   mStem,   dStem,   hStem   };
        var brs     = new[] { yBranch, mBranch, dBranch, hBranch };
        var labels  = new[] { "年", "月", "日", "時" };
        var stemSS  = new[] { yStemSS, mStemSS, "日主", hStemSS };
        var brSS    = new[] { yBranchSS, mBranchSS, dBranchSS, hBranchSS };
        var palaces = new[] {
            "祖先宮（早年0-16歲）", "父母宮（青年17-32歲）",
            "夫妻宮（自身伴侶）",   "子女宮（晚年）"
        };

        // pillarStars[pi] = 落在第 pi 柱的神煞（記錄來源柱與類型）
        // key = star name, value = "年支→支" / "日干→支" 等來源標注（取最先命中者）
        var pillarStars = new[] {
            new Dictionary<string, string>(), new Dictionary<string, string>(),
            new Dictionary<string, string>(), new Dictionary<string, string>()
        };

        // 四柱本身神煞：只以「年柱」(ki=0) 和「日柱」(ki=2) 為 KIND 基準查四柱
        // 月柱/時柱不作為基準（它們是被查對象，不是起算依據）
        foreach (int ki in new[] { 0, 2 })
        {
            if (DiZhiShenShaMap.TryGetValue(brs[ki], out var dzMap))
                for (int pi = 0; pi < 4; pi++)
                    if (dzMap.TryGetValue(brs[pi], out var ss))
                        foreach (var s in ss)
                            pillarStars[pi].TryAdd(s, $"{labels[ki]}支→支");

            if (TianGanShenShaMap.TryGetValue(stems[ki], out var tgMap))
                for (int pi = 0; pi < 4; pi++)
                    if (tgMap.TryGetValue(brs[pi], out var ss))
                        foreach (var s in ss)
                            pillarStars[pi].TryAdd(s, $"{labels[ki]}干→支");
        }

        var sb = new System.Text.StringBuilder();
        for (int i = 0; i < 4; i++)
        {
            sb.AppendLine($"▍{labels[i]}柱（{stems[i]}{brs[i]}）天干：{stemSS[i]}　地支：{brSS[i]}　宮位：{palaces[i]}");
            if (pillarStars[i].Count > 0)
            {
                var starList = pillarStars[i]
                    .Select(kv => {
                        var d = LfShenShaStarDesc(kv.Key);
                        string name = string.IsNullOrEmpty(d) ? kv.Key : $"{kv.Key}({d})";
                        return $"{name}〔{kv.Value}〕";
                    });
                sb.AppendLine($"  神煞：{string.Join("、", starList)}");
            }
            else
            {
                sb.AppendLine("  神煞：無");
            }
            if (i < 3) sb.AppendLine();
        }
        return sb.ToString().Trim();
    }

    // 空亡論斷（依文檔：空亡在八字中的用法.docx）
    // 節氣司令用事：計算出生日入節第幾天、當令藏干
    private static string LfCalcSiLing(int birthYear, int birthMonth, int birthDay, string mBranch, string dStem,
        CalendarDbContext? calDb = null)
    {
        if (birthYear <= 0 || birthMonth <= 0 || birthDay <= 0) return "";

        DateTime birthDate = new DateTime(birthYear, birthMonth, birthDay);
        string termName = "";
        DateTime termStart = birthDate;

        if (calDb != null)
        {
            // 直接查詢出生日的節氣欄位（DB 格式：「寒露」或「寒露第10天」）
            var birthEntry = calDb.CalendarEntries
                .Where(c => c.Year == birthDate.Year && c.SolarMonth == birthDate.Month && c.SolarDay == birthDate.Day)
                .FirstOrDefault();
            string rawSt = birthEntry?.SolarTerm ?? "";
            if (!string.IsNullOrEmpty(rawSt))
            {
                // 解析「寒露第10天」→ termName="寒露", dayInTerm=10
                var stMatch = System.Text.RegularExpressions.Regex.Match(rawSt, @"^(.+?)第(\d+)天$");
                if (stMatch.Success)
                {
                    termName  = stMatch.Groups[1].Value;
                    int dayN  = int.Parse(stMatch.Groups[2].Value);
                    termStart = birthDate.AddDays(-(dayN - 1));
                }
                else
                {
                    // 純節氣名稱（無天數）= 第1天
                    termName  = rawSt;
                    termStart = birthDate;
                }
            }
        }
        else
        {
            // 備援：使用 Ecan 掃描（月支對應的節氣名稱作為過濾依據）
            string expectedJie = LfBranchSolarTerms.TryGetValue(mBranch, out var jq) ? jq.jie : "";
            for (int offset = 0; offset <= 60; offset++)
            {
                var testDate = birthDate.AddDays(-offset);
                var testCal = new Ecan.EcanChineseCalendar(testDate);
                string st = testCal.SolarTermString;
                if (!string.IsNullOrEmpty(st) && st == expectedJie)
                {
                    termStart = testDate;
                    termName  = st;
                    break;
                }
            }
        }

        if (string.IsNullOrEmpty(termName)) return "";

        int dayInTerm = (int)(birthDate - termStart).TotalDays + 1; // 第1天起算

        // 查司令天干
        string siLingStem = "";
        string siLingPeriod = "";
        if (LfSilingTable.TryGetValue(mBranch, out var periods))
        {
            int prev = 0;
            foreach (var (s, end) in periods)
            {
                if (dayInTerm <= end)
                {
                    siLingStem = s;
                    siLingPeriod = $"第{prev + 1}-{end}天";
                    break;
                }
                prev = end;
            }
            if (string.IsNullOrEmpty(siLingStem))
            {
                siLingStem   = periods[^1].stem;
                siLingPeriod = $"第{periods[^2].endDay + 1}天以後";
            }
        }

        string siLingSS = string.IsNullOrEmpty(siLingStem) ? "" : LfStemShiShen(siLingStem, dStem);

        // 完整司令順序文字
        string fullOrder = "";
        if (LfSilingTable.TryGetValue(mBranch, out var pFull))
        {
            int p2 = 0;
            var parts = pFull.Select(pf => {
                string r = $"{pf.stem}（{p2 + 1}-{pf.endDay}天）";
                p2 = pf.endDay;
                return r;
            });
            fullOrder = string.Join(" → ", parts);
        }

        var sb = new StringBuilder();
        sb.AppendLine($"  入節第 {dayInTerm} 天（節氣：{termName}）");
        sb.AppendLine($"  當令藏干：{siLingStem}（{siLingPeriod}）→ 十神：{siLingSS}");
        sb.AppendLine($"  {mBranch}月司令順序：{fullOrder}");
        return sb.ToString().TrimEnd();
    }

    // 計算旬空二字（日柱天干基準）
    private static string[] LfCalcDayEmpty(string dStem, string dBranch)
    {
        var stems    = new[] { "甲","乙","丙","丁","戊","己","庚","辛","壬","癸" };
        var branches = new[] { "子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥" };
        int si = Array.IndexOf(stems, dStem);
        int bi = Array.IndexOf(branches, dBranch);
        if (si < 0 || bi < 0) return Array.Empty<string>();
        int start = (bi - si + 12) % 12;
        return new[] { branches[(start + 10) % 12], branches[(start + 11) % 12] };
    }

    // 年月時以日柱天干查旬空；日支以年柱天干查旬空
    private static string LfKongWang(
        string yStem, string yBranch, string mStem, string mBranch,
        string dStem, string dBranch, string hStem, string hBranch)
    {
        var stems   = new[] { "甲","乙","丙","丁","戊","己","庚","辛","壬","癸" };
        var branches = new[] { "子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥" };

        // 計算旬空：起始支 = (地支索引 - 天干索引 + 12) % 12；空亡 = 起始支+10、+11
        string[] CalcEmpty(string stem, string branch)
        {
            int si = Array.IndexOf(stems, stem);
            int bi = Array.IndexOf(branches, branch);
            if (si < 0 || bi < 0) return Array.Empty<string>();
            int start = (bi - si + 12) % 12;
            return new[] { branches[(start + 10) % 12], branches[(start + 11) % 12] };
        }

        // 日柱旬空 → 供年/月/時柱地支判斷
        var dayEmpty  = CalcEmpty(dStem, dBranch);
        // 年柱旬空 → 供日支判斷
        var yearEmpty = CalcEmpty(yStem, yBranch);

        if (dayEmpty.Length == 0 && yearEmpty.Length == 0) return "";

        bool yKong = dayEmpty.Contains(yBranch);
        bool mKong = dayEmpty.Contains(mBranch);
        bool dKong = yearEmpty.Contains(dBranch);  // 日支以年查
        bool hKong = dayEmpty.Contains(hBranch);

        if (!yKong && !mKong && !dKong && !hKong) return "";

        var sb = new System.Text.StringBuilder();
        sb.AppendLine($"（日柱旬空：{dayEmpty[0]}、{dayEmpty[1]}；年柱旬空：{yearEmpty[0]}、{yearEmpty[1]}）");
        sb.AppendLine();

        if (yKong)
            sb.AppendLine($"▍年柱（{yBranch}）落空亡：祖宗緣份薄，得不到祖上蔭庇或遺產，幼年至十六歲前多有阻滯。");
        if (mKong)
            sb.AppendLine($"▍月柱（{mBranch}）落空亡：父母緣薄，聚少離多，手足情分亦淡，青年運多有不順。");
        if (dKong)
            sb.AppendLine($"▍日支（{dBranch}）落空亡：夫妻緣薄，婚姻難以美滿，伴侶聚少離多，或同床異夢。");
        if (hKong)
            sb.AppendLine($"▍時柱（{hBranch}）落空亡：子女緣薄，或有損子之象，晚年易陷孤困之境。");

        // 坐空天干補充
        var seatNotes = new List<string>();
        if (yKong)  seatNotes.Add($"年干{yStem}");
        if (mKong)  seatNotes.Add($"月干{mStem}");
        if (dKong)  seatNotes.Add($"日干{dStem}");
        if (hKong)  seatNotes.Add($"時干{hStem}");
        if (seatNotes.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine($"坐空天干（{string.Join("、", seatNotes)}）力量虛浮，應事能力有所削減。");
        }

        return sb.ToString().Trim();
    }

    // 納音論斷（師傳.doc 七章）
    // yBranch_mBranch => 年生肖x月支; yBranch_dBranch => 年生肖x日支; yBranch_hBranch => 年生肖x時支
    private static string LfNaYin(string yStem, string yBranch, string mBranch,
                                   string dStem, string dBranch, string hBranch)
    {
        // 60甲子納音名稱
        var nayin60 = new Dictionary<string, string> {
            ["甲子"] = "海中金", ["乙丑"] = "海中金",
            ["丙寅"] = "爐中火", ["丁卯"] = "爐中火",
            ["戊辰"] = "大林木", ["己巳"] = "大林木",
            ["庚午"] = "路旁土", ["辛未"] = "路旁土",
            ["壬申"] = "劍鋒金", ["癸酉"] = "劍鋒金",
            ["甲戌"] = "山頭火", ["乙亥"] = "山頭火",
            ["丙子"] = "澗下水", ["丁丑"] = "澗下水",
            ["戊寅"] = "城頭土", ["己卯"] = "城頭土",
            ["庚辰"] = "白蠟金", ["辛巳"] = "白蠟金",
            ["壬午"] = "楊柳木", ["癸未"] = "楊柳木",
            ["甲申"] = "泉中水", ["乙酉"] = "泉中水",
            ["丙戌"] = "屋上土", ["丁亥"] = "屋上土",
            ["戊子"] = "霹靂火", ["己丑"] = "霹靂火",
            ["庚寅"] = "松柏木", ["辛卯"] = "松柏木",
            ["壬辰"] = "長流水", ["癸巳"] = "長流水",
            ["甲午"] = "沙中金", ["乙未"] = "沙中金",
            ["丙申"] = "山下火", ["丁酉"] = "山下火",
            ["戊戌"] = "平地木", ["己亥"] = "平地木",
            ["庚子"] = "壁上土", ["辛丑"] = "壁上土",
            ["壬寅"] = "金箔金", ["癸卯"] = "金箔金",
            ["甲辰"] = "覆燈火", ["乙巳"] = "覆燈火",
            ["丙午"] = "天河水", ["丁未"] = "天河水",
            ["戊申"] = "大驛土", ["己酉"] = "大驛土",
            ["庚戌"] = "釵釧金", ["辛亥"] = "釵釧金",
            ["壬子"] = "桑柘木", ["癸丑"] = "桑柘木",
            ["甲寅"] = "大溪水", ["乙卯"] = "大溪水",
            ["丙辰"] = "沙中土", ["丁巳"] = "沙中土",
            ["戊午"] = "天上火", ["己未"] = "天上火",
            ["庚申"] = "石榴木", ["辛酉"] = "石榴木",
            ["壬戌"] = "大海水", ["癸亥"] = "大海水",
        };

        // Ch.1 年柱納音五行論（命名原理）
        var nayinCh1 = new Dictionary<string, string> {
            ["城頭土"] = "以天干戊、己屬土，寅為艮山，土積為山，故曰城土也，天京玉壘，帝裡金城，龍蟠千時之形，虎踞四維之勢也。",
            ["壁上土"] = "丑雖是土家正位，而子則水旺之地，土見水多則為泥也，故曰壁上土也，氣屋開塞，物尚包藏，掩形遮體，內外不及故也。",
            ["大林木"] = "以辰為原野，巳為六陰，木至六陰則枝藏葉茂，以茂盛之大林木而生原野之間，故曰大林木，聲播九天，陰生萬頃。4、庚午辛未路旁土：以未中之土，生午木之旺火，則土旺於斯而受刑。土之所生未能自物，猶路旁土也，壯以及時乘原哉，木多不慮木。",
            ["大海水"] = "水冠帶在戍，臨官在亥，水則力厚矣。兼亥為江，非他水之比，故曰大海水。",
            ["大溪水"] = "寅為東旺維，卯為正東，水流正東，則其性順而川瀾池沼俱合而歸，故曰大溪水。注：維：方向。",
            ["大驛土"] = "申為坤，坤為地，酉為兌，兌為澤，戊己之土加於坤澤之上，非比他浮沉之土，故曰大驛土，氣以歸息，物當收斂，故雲。",
            ["天上火"] = "午為火旺之地，未、己之木又複生之，火性炎上故曰天上火。",
            ["天河水"] = "丙丁屬火，午為火旺之地，而納音乃水，水自火出，非銀漢而不能有也，故曰天河水，氣當升齊，沛然作霖，生旺有濟物之功。",
            ["屋上土"] = "以丙丁屬火，戍亥為天門，火既炎上則土非在下，故曰屋上土。提示：屋上土實際上應是磚瓦，戍亥一水一土，和而成泥，再加上火以燒烤，就成為磚瓦。修屋造房各有所用，既是屋上土，則需要有木的支撐和金的刻削裝點，屋上土方顯金碧輝煌，大富大貴之象。",
            ["山下火"] = "申為地戶，酉為日入之門，日至此時而無光亮，故曰山下火。提示：山下火實際夜晚的太陽，古人認為在夜晚太陽也和人一樣，在一個地方休息，因此遇土、遇木則吉，既是夜晚的陽光，自然不喜再見陽火，山頭火等。",
            ["山頭火"] = "以戍亥為天門，火照天門，其光至高，故曰山頭火也，天際斜暉，山頭落日散綺，因此返照舒霞，木白金光。",
            ["平地木"] = "戊為原野，亥為生木之地，夫木生於原野，則非一根一林之比，故曰平地木，惟貴雨露之功，不喜霜雪之積。",
            ["松柏木"] = "以木臨官在寅，帝旺在卯，木既生旺，則非柔弱之比，故曰松柏木也。積雪凝霜參天覆地，風撼笙簧，再余張旌施。",
            ["桑柘木"] = "子屬水而丑屬金，水剛生木而金則伐之，猶桑柘方生而人便以戕伐。故曰桑柘木。",
            ["楊柳木"] = "以木死于午墓於未，木既死墓，唯得天干壬癸之水以生之，終是柔木，故曰楊柳木，萬縷不蠶之絲，千條不了之線。",
            ["沙中土"] = "土庫在辰，而絕在巳，而天本丙丁之火至辰為冠帶，而臨官在巳，土既張絕，得旺火生之而復興，故曰沙中土。",
            ["泉中水"] = "金既臨官在申，帝旺在酉，旺則生自以火，然方生之際方量未興，故曰井泉水也。氣息在靜，過而不竭，出而不窮。注．旺則生自以火：此指水的來源，五行論認為金生水，金要變成水需要有火，就如金屬用火熔成水一樣。",
            ["海中金"] = "以子為水，又為湖，又為水旺之地，兼金死於子、墓于丑，水旺而金死、墓，故曰海中之金，又曰氣在包藏，使極則沉潛。子，五行是水，是湖泊之水，是水勢旺盛的地方，在五行中金死在子而墓在丑，水旺金死、墓，尤如大海中之金子故曰大海金。",
            ["澗下水"] = "以水旺于子衰於丑，旺而反衰則不能成江河，故曰澗下水，出環細浪，雪湧飛淌，匯流三峽之傾瀾，望下尋之倒。",
            ["爐中火"] = "以寅為三陽，卯為四陰，火既得位，又得寅卯之木以生之，此時天地開爐，萬物始生，故曰爐中火，天地為爐，陰陽為炭。3、戊辰己巳大林木：以辰為原野，巳為六陰，木至六陰則枝藏葉茂，以茂盛之大林木而生原野之間，故曰大林木，聲播九天，陰生萬頃。",
            ["白蠟金"] = "以金養於辰而生於巳，形質初成，未能堅利，故曰白蠟金，氣漸發生，交棲日月之光，凝陰陽之氣。",
            ["石榴木"] = "申為七月，酉為八月，此時木則絕矣，惟石榴之木複實，故曰石榴木；氣歸靜肅，物漸成實，木居金生其味，秋果成實矣。",
            ["覆燈火"] = "辰為日時，巳為日之將午，豔陽之勢光於天下，故曰覆燈火，金盞搖光，玉台吐豔，照日月不照處，明天下未明時。",
            ["路旁土"] = "以未中之土，生午木之旺火，則土旺於斯而受刑。土之所生未能自物，猶路旁土也，壯以及時乘原哉，木多不慮木。",
            ["金箔金"] = "寅卯為木旺之地，木旺則金贏，且金絕於寅，胎于卯，金既無力，故曰金箔金。",
            ["釵釧金"] = "金至戍而衰，至亥而病，全既衰病則誠柔矣，故曰釵釧金，形已成器，華飾光豔乎？生旺者乎？戳體火盛：戳體火盛的盛，這裡指器皿，即盛金之模。",
            ["長流水"] = "辰為水庫，巳為金的長生之地，金則生水，水性已存，以庫水而逢生金，泉源終不竭，故曰長流水也。勢居東南，貴安靜。",
            ["霹靂火"] = "丑屬土，子屬水，水居正位而納音乃火，水中之火非龍神則無，故曰霹靂火，電擊金蛇之勢，雲馳鐵騎之奔，變化之象。",
        };

        // Ch.2 日柱納音細注（六十甲子完整，含神煞名稱）
        var nayinCh2 = new Dictionary<string, string> {
            ["甲子"] = "甲子金，為寶物，喜金木旺地。進神喜，福星，平頭，懸針，破字。",
            ["乙丑"] = "乙丑金，為頑礦，喜火及南方日時。福星，華蓋，正印。",
            ["丙寅"] = "丙寅火，為爐炭，喜冬及木。福星，祿刑，平頭，聾啞。",
            ["丁卯"] = "丁卯火，為爐煙，喜巽地及秋冬。平頭，截路，懸針。",
            ["戊辰"] = "戊辰木，山林山野處不材之木，喜水。祿庫，華蓋，水馬庫，棒杖，伏神，平頭。",
            ["己巳"] = "己巳木，山頭花草，喜春及秋。祿庫，八吉，闕字，曲腳。",
            ["庚午"] = "庚午土，路旁幹土，喜水及春。福星，官貴，截路，棒杖，懸針。",
            ["辛未"] = "辛未土，含萬寶，待秋成，喜秋及火。華蓋，懸針，破字。",
            ["壬申"] = "壬申金，戈戟，大喜子午卯酉。平頭，大敗，妨害，聾啞，破字，懸針。",
            ["癸酉"] = "癸酉金，金之椎鑿，喜木及寅卯。伏神，破字，聾啞。",
            ["甲戌"] = "甲戌火，火所宿處，喜春及夏。正印，華蓋，平頭，懸針，破字，棒杖。",
            ["乙亥"] = "乙亥火，火之熱氣，喜土及夏。天德，曲腳。",
            ["丙子"] = "丙子水，江湖，喜木及土。福星，官貴，平頭，聾啞，交神，飛刃。",
            ["丁丑"] = "丁丑水，水之不流清澈處，喜金及夏。華蓋，進神，平頭，飛刃，闕字。",
            ["戊寅"] = "戊寅土，堤阜城郭，喜木及火。伏神，棒杖，聾啞。",
            ["己卯"] = "己卯土，破堤敗城，喜申酉及火。進神，短夭，九丑，闕字，曲腳，懸針。",
            ["庚辰"] = "庚辰金，錫蠟，喜秋及微木。華蓋，大敗，棒杖，平頭。",
            ["辛巳"] = "辛巳金，金之幽者，雜沙石，喜火及秋。天德，福星，官貴，截路，大敗，懸針，曲腳。",
            ["壬午"] = "壬午木，楊柳幹節，喜春夏。官貴，九丑，飛刃，平頭，聾啞，懸針。",
            ["癸未"] = "癸未木，楊柳根，喜冬及水，亦宜春。正印，華蓋，短夭，伏神，飛刃，破字。",
            ["甲申"] = "甲申水，甘井，喜春及夏。破祿馬，截路，平頭，破字，懸針。",
            ["乙酉"] = "乙酉水，陰壑水，喜東方及南。破祿，短夭，九丑，曲腳，破字，聾啞。",
            ["丙戌"] = "丙戌土，堆阜，喜春夏及水。天德，華蓋，平頭，聾啞。",
            ["丁亥"] = "丁亥土，平原，喜火及木。天乙，福星，官貴，德合，平頭。",
            ["戊子"] = "戊子火，雷也。喜水及春夏，得土而神天。伏神，短夭，九丑，杖刑，飛刃。",
            ["己丑"] = "己丑火，電也，喜水及春夏，得地而晦。華蓋，大敗，飛刃，曲腳，闕字。",
            ["庚寅"] = "庚寅木，松柏幹節，喜秋冬。破祿馬，相刑，杖刑，聾啞。",
            ["辛卯"] = "辛卯木，松柏之根，喜水土及春。破祿，交神，九丑，懸針。",
            ["壬辰"] = "壬辰水，龍水，喜雷電及春夏。正印，天德，水祿馬庫，退神，平頭，聾啞。",
            ["癸巳"] = "癸巳水，水之不息，流入海，喜亥子，乃變化。天乙，官貴，德合，伏馬，破字，曲腳。",
            ["甲午"] = "甲午金，百煉精金，喜水木土。進神，德合，平頭，破字，懸針。",
            ["乙未"] = "乙未金，爐炭余金，喜大火及土。華蓋，截路，曲腳，破字。",
            ["丙申"] = "丙申火，白茅野燒，喜秋冬及木。平頭，聾啞，大敗，破字，懸針。",
            ["丁酉"] = "丁酉火，鬼神之靈響，火無形者，喜辰戌丑未。天乙，喜神，平頭，破字，聾啞，大敗。",
            ["戊戌"] = "戊戌木，蒿艾之枯者，喜火及春夏。華蓋，大敗，八專，杖刑，截路。",
            ["己亥"] = "己亥木，蒿艾之茅，喜水及春夏。闕字，曲腳。",
            ["庚子"] = "庚子土，土中空者，屋宇也，喜木及金。木德合，杖刑。",
            ["辛丑"] = "辛丑土，墳墓，喜木及火與春。華蓋，懸針，闕字。",
            ["壬寅"] = "壬寅金，金之華飾者，喜木及微火。截路，平頭，聾啞。",
            ["癸卯"] = "癸卯金，環鈕鈐鐸，喜盛火及秋。貴人，破字，懸針。",
            ["甲辰"] = "甲辰火，燈也，喜夜及水，惡晝。華蓋，大敗，平頭，破字，懸針。",
            ["乙巳"] = "乙巳火，燈光也，尤喜申酉及秋。正祿馬，大敗，曲腳，闕字。",
            ["丙午"] = "丙午火，月輪，喜夜及秋，水旺也。喜神，羊刃，交神，平頭，聾啞，懸針。",
            ["丁未"] = "丁未水，火光也，喜辰戌丑未。華蓋，羊刃，退神，八專，平頭，破字。",
            ["戊申"] = "戊申土，秋間田地，喜申酉及火。福星，伏馬，杖刑，破字，懸針。",
            ["己酉"] = "己酉土，秋間禾稼，喜申酉及冬。退神，截路，九丑，闕字，曲腳，破字，聾啞。",
            ["庚戌"] = "庚戌金，刃劍之餘，喜微火及木。華蓋，杖刑。",
            ["辛亥"] = "辛亥金，鍾鼎實物，喜木火及土。正祿馬，懸針。",
            ["壬子"] = "壬子木，傷水多之木，喜火土及夏。羊刃，九丑，平頭，聾啞。",
            ["癸丑"] = "癸丑木，傷水少之木，喜金水及秋。華蓋，福星，八專，破字，闕字，羊刃。",
            ["甲寅"] = "甲寅水，雨也，喜夏及火。正祿馬，福神，八專，平頭，破字，懸針，聾啞。",
            ["乙卯"] = "乙卯水，露也，喜水及火。建祿，喜神，八專，九刃，曲腳，懸針。",
            ["丙辰"] = "丙辰土，堤岸，喜金及木。祿庫，正印，華蓋，截路，平頭，聾啞。",
            ["丁巳"] = "丁巳土，土之沮，喜火及西北。祿庫，平頭，闕字，曲腳。",
            ["戊午"] = "戊午火，日輪，夏則人畏，冬則人愛，忌戊子、己丑、甲寅、乙卯。伏神，羊刃，九丑，棒杖，懸針。",
            ["己未"] = "己未火，日光，忌夜。福星，華蓋，羊刃，闕字，曲腳，破字。",
            ["庚申"] = "庚申木，榴花，喜夏，不宜秋冬。建祿馬，八專，杖刑，破字，懸針。",
            ["辛酉"] = "辛酉木，榴子，喜秋及夏。建祿，交神，九丑，八專，懸針，聾啞。",
            ["壬戌"] = "壬戌水，海也，喜春夏及木。華蓋，退神，平頭，聾啞，杖刑。",
            ["癸亥"] = "癸亥水，百川，喜金土火。伏馬，大敗，破字，截路。",
        };

        // Ch.3 年命推算（命運走向）
        var nayinCh3 = new Dictionary<string, string> {
            ["劍鋒金"] = "日時遇長流在壬辰為寶劍化為青龍，癸巳亦得，此劍不能通變，然癸丑為劍氣動鬥最吉，松柏楊柳亦吉但多聚散，大林平地嫌有土制主勞苦。火見神龍陰陽交遇，如壬申逢己丑癸酉逢戊子方為上格遇天上爐中二火無水救則夭，諸土見皆不吉，以其埋沒。",
            ["城頭土"] = "此土有成未成作兩般論，凡遇見路傍為巳成之土不必用火，若無路傍為未成之土，須用火。大都城土皆須資木楊柳癸未最佳，壬午則忌，桑柘癸丑為上，壬子次之。庚寅辛卯就位相克，則城崩不甯何以安人也。耶如見木無夾輔只以貴人祿馬論之。",
            ["壁上土"] = "此乃人間壁土，非平地何以為*，子午天地正柱逢之尤為吉慶。凡見木皆可為主，庚寅辛卯亦是棟樑，只辛酉衝破子卯相刑，大林有風若無承載之土加凡水主作事難成，貧賤而夭。",
            ["大林木"] = "此木生居東南春夏之交，長養成林，全假艮土為源癸丑為山，三命無破陷最為福厚權貴。戊辰為上己巳次之，土遇路傍為負載，戊辰見辛未為貴，己巳見庚午為祿主福，壁屋二土再得劍金，則大林之木取為棟樑，成格最吉。",
            ["大海水"] = "壬戍有土為濁，癸亥干支純水納音又水為清。壬戍人嫌山以土氣太盛，有金清之方吉，癸亥最喜見山然後海水之性始安，澗下丁丑為山天河與海上下相通，柱中有木為槎，則乘槎入天河極為上格。長流大溪等水畢竟皆歸於大海，以海水不擇細流故能成其大也。",
            ["大溪水"] = "井泉靜而止，澗下有丑為艮，天河沾潤，大海朝宗此四水皆吉。長流有風，獨不宜見，此水以清金為助養，惟釵金最宜，蠟金亦清，若有釵金對沖，則不宜。海中雖無造化甲子屬坎，乙丑為艮乃歸源之地亦吉。泊金最微不能相生，豈有超顯之理。",
            ["大驛土"] = "天河丙午而己酉得之，丁未而戊申得之亦為貴祿主福長流。戊申見癸巳，己酉見壬辰亦吉，多逢則不寧靜，大溪乙卯為東震發生之義單見亦吉。海水對穿土不能勝，日時遇主夭，得山稍輕，內戊申見癸亥戊癸合，申亥為地天交泰反吉。",
            ["天上火"] = "此火見木謂之震折，要日時有風興水方得。大林木有辰巳松柏石榴有卯酉惟此三木主貴。午見木多猶可，未三四木勞苦之命也。見金且能照耀不能克濟。釵金有戍亥泊金有寅卯主吉。劍金為耀日月之光必主少年及第。佘金則殃。",
            ["天河水"] = "土不能克，故見土不忌，而且有滋潤之益，天上之水地金難生，故見金難益，而亦有涵秀之情。生旺太過則為淫潦反傷於物，死絕太多則為旱乾，又不能生物，要生三秋得時為貴也。水喜長流大海，內丙午宜癸巳癸亥，丁未宜壬辰壬戍，陰陽互見尤吉。",
            ["屋上土"] = "不宜見火，爐中丙寅最凶丁卯稍可，太陽霹靂可取相資，山下山頭有木生之則禍。燈火丙戍見乙巳為上，丁卯見甲辰次之謂之火土入堂格。若柱中木多，亦不為吉，水宜天河澗下井泉皆吉，如先得平地木成格大貴。",
            ["山下火"] = "故妙選有螢火照水格，遇秋生則貴為卿監，是以此火喜水地支逢亥子或納音水更遇申酉月是也，或以山下之火最喜木與山更得風來增輝為貴又不以螢火論矣。大林有辰巳為風桑柘木有癸丑為山松柏平地最吉更得風助主貴，若風多吹散主夭。",
            ["山頭火"] = "大既宜山木與風，木喜大林與松柏以辰巳有風，寅卯歸祿，更得癸丑為上土木主貴，無山則木無所依火無所見，縱有風亦不光顯佘木無用，只以祿馬看，水宜澗下名為交泰主吉。井泉清水有木助之亦吉大溪甲戍見甲寅乙亥見乙卯卻真祿俱吉。",
            ["平地木"] = "不喜雪霜之積，此乃地上之茂材，人間之屋木，戊戍為棟己亥為梁最宜互換見之，須以土為基，土愛路傍為正格，更逢子午尤貴，以子午為天地正柱故也。屋壁城頭三土以此木相資中間有升化尤吉，砂驛無用，日時見之主災夭。",
            ["松柏木"] = "命中有火最忌爐中就位相生，再加風木灰飛煙滅，五行無水主夭折，山頭山下太陽佛燈皆不可犯，寅人尤忌。戊午丙寅以木不南奔寅午三合火局故也。辛卯無害，霹火雖可滋生運加凡火主凶。土見路傍似無足貴若無死木其福還真。",
            ["桑柘木"] = "最愛砂土以為根基又以辰巳為蠶食之地不宜刑沖互破，路傍大驛二土次吉佘土無益。水喜天河為雨露之滋，長流溪澗井泉諸水皆可相依，亦須先得土為基更加祿貴為妙，淪海水漂泛無定，無土主夭。見火，燈頭最吉亦以辰巳之中為蠶位故也。",
            ["楊柳木"] = "午未木之死墓壬癸木之滋潤，此木根基惟喜砂土，見艮則依倚搖金，遇寅卯則東方得地，辛丑有山庚子不如，戊寅雖吉己卯尤勝，丙辰丁巳卻嫌戍亥對沖，若見大驛有丑為山邊之驛稍可，無土獨見此土主夭賤。",
            ["海中金"] = "以甲子見癸亥，是不用火，逢空有蚌珠照月格。以甲子見己未是欲合化互貴，蓋以海金無形，非空動則不能出現。而乙丑金庫，非旺火則不能陶鑄故也。如甲子見戊寅、庚午是土生金，乙丑見丙寅、丁卯是火制金，又天干逢三奇，此等格局，無有不貴。",
            ["澗下水"] = "佘金以祿貴參之，取其資生，惡其衝破，見木一位無妨，二三則主勞苦，亦以貴人祿馬參之，命中見土主人多濁，天元若木或化水則主清吉。砂中屋上二土其氣稍清，路傍大驛則濁甚，主財散禍生，如辰戍丑未土局，其凶尤，以水濁土混故也。",
            ["爐中火"] = "此火炎上，喜得木生，平地之木為上以丙寅見己亥謂之天乙貴見戊戍謂之歸庫故吉，丁卯次之。然丙丁火自生無木庶幾，丁卯火自敗若無木則凶。且此火以金為用，更得金方應化機，但丁卯無木而遇金主勞苦之命。夫寅見木多火炎而無水制則主夭，卯見三四木不妨。",
            ["白蠟金"] = "此金惟喜火煉，需爐中炎火。然庚辰見之，若無水濟，主貧夭，辛巳卻以貴論，緣巳金生之地，見丙寅化水逢貴故也。山下火生早，主榮貴，亦需水助方得。井泉、大溪俱為貴格，庚官在丁，辛官在丙，故庚見丁丑，官貴俱全；辛見丙子，不如癸巳更清，不貴即富。",
            ["石榴木"] = "路壁驛砂四土有山助之亦吉，若無何用，見砂中最吉，泊金干支有水木而納音金，榴木干支金而納音木皆脫去本性而互換歸旺，以木旺寅卯金旺申酉各得其位謂之功侔造化格主大貴。",
            ["釵釧金"] = "海水貧夭，天河辛亥見之無妨，丙午真火庚戍湊成火局有傷此金也，太陽火，日生顯耀。佛燈火夜間顯耀皆宜見。但甲辰乙巳與庚戍辛亥相沖，陰陽互見為妙。戊子己丑與丙午丁未相持二火忌疊見，非貧則夭。爐中火庚戍最忌，辛亥見之有丙辛合化水稍吉。",
            ["長流水"] = "逢木雖漂泛，而桑柘癸丑為山，楊柳癸未為圓，年時得此水圍之為水繞花堤大貴格。松柏石榴天干有金相生，大林平地雖嫌有土，而癸巳見戊辰，壬辰見己亥俱吉。水為同類，澗下則丁壬合化，天上則雨露相資。井泉大溪無用。",
            ["霹靂火"] = "此火須資風雷方為變化，若五行得此一件皆主享通。如日時見大海癸亥引凡入聖己丑為上戊子次之。見大溪乙為雷火變化己丑吉戊子次之。辰巳為風運中遇之尤佳。天上水名為既濟主吉遇之者稟性含靈聰明特異。長流無用。澗下雖就位相克，此神火也是不忌，有風亦顯。",
        };

        // Ch.4 論青紅（生肖+納音特質）
        var nayinCh4 = new Dictionary<string, string> {
            ["丁丑"] = "湖內之牛。待人和睦，衣祿不虧。早年財不豐富，晚景有餘，骨肉無情，見遲方好；夫婦和順，女命旺夫，持家，賢良之命也。城牆土命，屬相為虎：過山之虎。為人性急猛烈，易惡，反眼無情，早年勤儉，離祖方能發達，必主聰明伶俐。城牆土命，屬相為兔：出林之兔。",
            ["丁亥"] = "過山之豬。性溫，手巧聰明，自立家業。兒女有刑傷，見遲方好，一生好善；女命衣祿平穩，無風浪。霹雷火命，屬相為鼠：倉內之鼠。計算超人，文武皆能，兒女早年有克，遲得安全，夫婦白頭和好；女命賢良發達。霹雷火命，屬相為牛：欄內之牛。",
            ["丁卯"] = "為望月之兔，一生手足不停，身心不閑，衣祿可以，可算富貴，女命好靜，一生有福之命。《窮大林木命，屬相為龍：為清溫之龍，喜氣春風，出入壓群。利官，近貴，骨肉刑傷，兒女無力；女命主溫良賢達，心直口快，必嫁好夫。",
            ["丁巳"] = "塘內之蛇。利官有貴，性剛不順，身居長位，事業顯榮；女命美麗，賢達會治家，有財旺夫之命天上火命，屬相為馬：殿內之馬。衣祿自然有，五官端正，志氣寬宏，少年有災，骨肉有刑；女命難得姐妹之助，後來興旺自立之命。天上火命，屬相為羊：草野之羊。",
            ["丁未"] = "失群之羊。一生喜怒無常，多口舌，是非不可無。有名利，骨肉生疏，妻遲子晚好；女命財旺助夫，利夫小之命。大驛土命，屬相為猴：獨立之猴。",
            ["丁酉"] = "樓宿之雞。說話流利，巧言善辯，辦事公正，無私心，衣良平穩。六親無緣，晚景旺相，女命益夫助夫興家。山頭火命，屬相為狗：守門之犬。口快舌尖，用心不用力，有權柄，名聲四海；夫婦相合；女命旺夫生財，興家立業。35亥山頭火命，屬相為豬：過往之豬。",
            ["丙午"] = "路行之馬。一生清閒，初年有財難聚，無主意，兄弟各自成家，後運旺財發家；女命清正手巧持家。天河水命，屬相為羊：失群之羊。一生喜怒無常，多口舌，是非不可無。有名利，骨肉生疏，妻遲子晚好；女命財旺助夫，利夫小之命。大驛土命，屬相為猴：獨立之猴。",
            ["丙子"] = "田內之鼠。為人膽大有權，機謀超人，早年平，中年好轉，老年大好。女命主口舌利快。澗下水命，屬相為牛：湖內之牛。待人和睦，衣祿不虧。早年財不豐富，晚景有餘，骨肉無情，見遲方好；夫婦和順，女命旺夫，持家，賢良之命也。城牆土命，屬相為虎：過山之虎。",
            ["丙寅"] = "為山林之虎。其性格不定，心直口快，手足忙碌，不利官。女主賢良，聰明伶俐之命。爐中火命，屬相為兔：為望月之兔，一生手足不停，身心不閑，衣祿可以，可算富貴，女命好靜，一生有福之命。《窮大林木命，屬相為龍：為清溫之龍，喜氣春風，出入壓群。",
            ["丙戌"] = "白眼之犬。一生多吉利，有招財進寶象徵，自立家業，至老榮華；女命財旺，興家有方。屋上土命，屬相為豬：過山之豬。性溫，手巧聰明，自立家業。兒女有刑傷，見遲方好，一生好善；女命衣祿平穩，無風浪。霹雷火命，屬相為鼠：倉內之鼠。",
            ["丙申"] = "清秀之猴，情高手巧，聰明伶俐，有智謀，遇事多變。和氣春風，功名有分，有賢德之妻，女主貌美色姿。33酉劍鋒金命，屬相為雞：樓宿之雞。說話流利，巧言善辯，辦事公正，無私心，衣良平穩。六親無緣，晚景旺相，女命益夫助夫興家。",
            ["丙辰"] = "上天之龍。一生衣食無虧，有名有利，四海春風，愛友喜朋，女命主賢能手巧口巧，旺夫興家之命。沙中土命，屬相為蛇：塘內之蛇。利官有貴，性剛不順，身居長位，事業顯榮；女命美麗，賢達會治家，有財旺夫之命天上火命，屬相為馬：殿內之馬。",
            ["乙丑"] = "海中之牛，少年運氣不好，有父母重婚之兆。夫婦皆好，兒女不孤，愛好九流藝術。爐中火命，屬相為虎：為山林之虎。其性格不定，心直口快，手足忙碌，不利官。女主賢良，聰明伶俐之命。",
            ["乙亥"] = "過往之豬。與人和順，初年多災。父母有刑傷，再拜無妨，夫婦無刑皆老，存心中正，豐足興旺，子息有克破，晚兒方好。澗下水命，屬相為鼠：田內之鼠。為人膽大有權，機謀超人，早年平，中年好轉，老年大好。女命主口舌利快。澗下水命，屬相為牛：湖內之牛。",
            ["乙卯"] = "得道之兔。志高氣昂，多計多謀，進事稱心如意，文武皆能。女命主一生福壽無虧之命。沙中土命，屬相為龍：上天之龍。一生衣食無虧，有名有利，四海春風，愛友喜朋，女命主賢能手巧口巧，旺夫興家之命。沙中土命，屬相為蛇：塘內之蛇。",
            ["乙巳"] = "福氣之蛇，有功名之徵兆，伶俐聰明，夫婦命和，做事能遂心願，財豐田足；女命衣良無缺，良婦之命。30午路旁土命，屬相為馬：福氣之蛇，有功名之徵兆，伶俐聰明，夫婦命和，做事能遂心願，財豐田足；女命衣良無缺，良婦之命。",
            ["乙未"] = "敬重之羊。容貌端正，少年勤儉，初年平順無大難，晚年聚財巨富，兄弟少力，子息不孤，立家興業之命。山頭火命，屬相為狗：守門之犬。口快舌尖，用心不用力，有權柄，名聲四海；夫婦相合；女命旺夫生財，興家立業。山頭火命，數相為豬：過往之豬。",
            ["乙酉"] = "唱午之雞。心直口快，有口無噁心，志氣高，衣食足，福壽雙全，弟兄多，但不得力；女命主興財，平穩之命。屋上土命，屬相為狗：白眼之犬。一生多吉利，有招財進寶象徵，自立家業，至老榮華；女命財旺，興家有方。屋上土命，屬相為豬：過山之豬。",
            ["壬午"] = "軍中之馬。為人勤儉，父母有刑傷，災危可折，早年財物有緊，晚年旺相有餘；女命，主興家也，為賢良之婦。43未楊柳木命，屬相為羊：群內之羊。一生有快樂之心，廣行方便，只是救人無恩，反招是非，命中財來不聚，女命賢德持家，晚年榮昌。",
            ["壬子"] = "屋上之鼠，為人少學少成，心情急躁，辦事有始無終，幼年有災，兄弟少力，男主女大，女主夫長，可配夫妻美滿。聰明、善良之命。25丑海中金命，屬相為牛：海中之牛，少年運氣不好，有父母重婚之兆。夫婦皆好，兒女不孤，愛好九流藝術。",
            ["壬寅"] = "過林之虎。心直口快，不暗地生非。夫婦，兒女有克，遲方好。初年財不聚，事事難好，晚景豐足之命63金箔金命，屬相為兔：樹木之兔。酒食不欠，福壽有餘，逢凶化吉，一生平安、興隆一世；女命中末運氣欠佳，操持家務一生。佛燈火命，屬相為龍：有道之龍。",
            ["壬戌"] = "雇家之犬。一生好行善，初年平平，末年興旺。",
            ["壬申"] = "清秀之猴，情高手巧，聰明伶俐，有智謀，遇事多變。和氣春風，功名有分，有賢德之妻，女主貌美色姿。33酉劍鋒金命，屬相為雞：樓宿之雞。說話流利，巧言善辯，辦事公正，無私心，衣良平穩。六親無緣，晚景旺相，女命益夫助夫興家。",
            ["壬辰"] = "行雨之龍。為人辛勤勞碌，終日手足不停，早歲運悔，晚景財旺發福；女命為一生操持家業，助夫興家。53巳長流水命，屬相為蛇：草中之蛇。聰明伶俐，財來不聚，晚年不同前，發財有福；女命賢良，興家助夫。沙中金命，屬相為馬：春中之馬。",
            ["己丑"] = "海中之牛，少年運氣不好，有父母重婚之兆。夫婦皆好，兒女不孤，愛好九流藝術。爐中火命，屬相為虎：為山林之虎。其性格不定，心直口快，手足忙碌，不利官。女主賢良，聰明伶俐之命。",
            ["己亥"] = "道院之豬。手巧技藝多，一生衣祿安穩，六親有克，夫婦各順；女命主清閒，老來樂觀之命。60壁上土命，屬相為鼠：梁上之鼠。人人尊重，妻主安穩清正，持家有權柄，衣食不虧，通達之命；女命主旺財，並逢凶化吉。61丑壁上土命，屬相為牛：路途之牛。",
            ["己卯"] = "為望月之兔，一生手足不停，身心不閑，衣祿可以，可算富貴，女命好靜，一生有福之命。《窮大林木命，屬相為龍：為清溫之龍，喜氣春風，出入壓群。利官，近貴，骨肉刑傷，兒女無力；女命主溫良賢達，心直口快，必嫁好夫。",
            ["己巳"] = "福氣之蛇，有功名之徵兆，伶俐聰明，夫婦命和，做事能遂心願，財豐田足；女命衣良無缺，良婦之命。30午路旁土命，屬相為馬：福氣之蛇，有功名之徵兆，伶俐聰明，夫婦命和，做事能遂心願，財豐田足；女命衣良無缺，良婦之命。",
            ["己未"] = "草野之羊。口快善變，衣食自來，前程顯達，有貴人扶，百事華榮，旺相之命。80申石榴木命，屬相為猴：食果之猴。一生手足不停，清高利官，但救人不得好報，女命六親冷淡老來興旺之命。81酉石榴木命，屬相為雞：籠藏之雞。",
            ["己酉"] = "報曉之雞。一生衣食有餘，但六親難靠，早子見婦，聰明好學，百一皆通；女命淫邪多變，一生無災之命。70戌釵釧金命，屬相為狗：寺院之犬。一生快活，丑年有災，作事機巧，百事如意；女命主聰明伶俐，手巧，能助夫興家旺財之命。",
            ["庚午"] = "福氣之蛇，有功名之徵兆，伶俐聰明，夫婦命和，做事能遂心願，財豐田足；女命衣良無缺，良婦之命。31未路旁土命，屬相為羊：得祿之羊，一生情性寬大志高，少年不利，夫婦和諧，頭胎見子有刑，女命持家興旺。",
            ["庚子"] = "梁上之鼠。人人尊重，妻主安穩清正，持家有權柄，衣食不虧，通達之命；女命主旺財，並逢凶化吉。61丑壁上土命，屬相為牛：路途之牛。對人溫和，初年有意外之危。衣食雖有，但進財不存，老來福壽雙至；女命賢良，旺財之命。金箔金命，屬相為虎：過林之虎。",
            ["庚寅"] = "出山之虎。心情急躁，口噁心善，有志氣，握權柄，利官貴。反復多變，衣食不虧；女命榮華，內助興家。51松柏木命，屬相為兔：蟾窟之兔。口快，有權。利官近貴身閑，心忙，六親少利，自立自成家業；女命持家，興家旺相。長流水命，屬相為龍：行雨之龍。",
            ["庚戌"] = "寺院之犬。一生快活，丑年有災，作事機巧，百事如意；女命主聰明伶俐，手巧，能助夫興家旺財之命。71亥釵釧金命，屬相為豬：圈內之豬。一生不貪閒事，不惹是非，早年財物不足，到老榮華；女命主體壯，能興家立業桑松木命，屬相為鼠：葉上之鼠。",
            ["庚申"] = "食果之猴。一生手足不停，清高利官，但救人不得好報，女命六親冷淡老來興旺之命。81酉石榴木命，屬相為雞：籠藏之雞。一生伶俐，精神爽朗，口舌能辯，六親生疏，女命主賢德治家，老來榮華之命。大海水命，屬相為狗：雇家之犬。",
            ["庚辰"] = "怒性之龍。一生和氣欠佳，勞碌雪霜，不受人欺，利官近貴，名利雙全，衣食有餘之命。41巳白蠟金命，屬相為蛇：冬藏之蛇。為人機謀多變，志氣人，一生衣食無虧，主有貴人扶持之徵兆，晚年快樂之命。楊柳木命，屬相為馬：軍中之馬。",
            ["戊午"] = "殿內之馬。衣祿自然有，五官端正，志氣寬宏，少年有災，骨肉有刑；女命難得姐妹之助，後來興旺自立之命。天上火命，屬相為羊：草野之羊。口快善變，衣食自來，前程顯達，有貴人扶，百事華榮，旺相之命。80申石榴木命，屬相為猴：食果之猴。",
            ["戊子"] = "屋上之鼠，為人少學少成，心情急躁，辦事有始無終，幼年有災，兄弟少力，男主女大，女主夫長，可配夫妻美滿。聰明、善良之命。25丑海中金命，屬相為牛：海中之牛，少年運氣不好，有父母重婚之兆。夫婦皆好，兒女不孤，愛好九流藝術。",
            ["戊寅"] = "為山林之虎。其性格不定，心直口快，手足忙碌，不利官。女主賢良，聰明伶俐之命。爐中火命，屬相為兔：為望月之兔，一生手足不停，身心不閑，衣祿可以，可算富貴，女命好靜，一生有福之命。《窮大林木命，屬相為龍：為清溫之龍，喜氣春風，出入壓群。",
            ["戊戌"] = "進山之犬。和氣待人，家業自創自立，早年不幸，物財耗散，晚來宜得財，行師藝術，中平之命。平地木命，屬相為豬：道院之豬。手巧技藝多，一生衣祿安穩，六親有克，夫婦各順；女命主清閒，老來樂觀之命。60壁上土命，屬相為鼠：梁上之鼠。",
            ["戊申"] = "獨立之猴。一生急性，作事反復無常，勞碌辛苦，但利官近貴，兒女刑傷，財帛足用；女命賢良懂事，文武皆精，手巧心活，利子女之命。大驛土命，屬相為雞：報曉之雞。一生衣食有餘，但六親難靠，早子見婦，聰明好學，百一皆通；女命淫邪多變，一生無災之命。",
            ["戊辰"] = "為清溫之龍，喜氣春風，出入壓群。利官，近貴，骨肉刑傷，兒女無力；女命主溫良賢達，心直口快，必嫁好夫。大林木命，屬相為蛇：福氣之蛇，有功名之徵兆，伶俐聰明，夫婦命和，做事能遂心願，財豐田足；女命衣良無缺，良婦之命。",
            ["甲午"] = "春中之馬。對人和氣如春風，好交朋友，利官近貴。遇凶化吉，一生無刑。女命能言善語，口尖舌快。55未沙中金命，屬相為羊：敬重之羊。容貌端正，少年勤儉，初年平順無大難，晚年聚財巨富，兄弟少力，子息不孤，立家興業之命。山頭火命，屬相為狗：守門之犬。",
            ["甲子"] = "屋上之鼠，為人少學少成，心情急躁，辦事有始無終，幼年有災，兄弟少力，男主女大，女主夫長，可配夫妻美滿。聰明、善良之命。25丑海中金命，屬相為牛：海中之牛，少年運氣不好，有父母重婚之兆。夫婦皆好，兒女不孤，愛好九流藝術。",
            ["甲寅"] = "立定之虎。一生利官利貴，家豐財，足衣食，財源繁多，有刑父母，雙親有重拜之災；女主夫大，男主妻長之命。75大溪水命，屬相為兔：得道之兔。志高氣昂，多計多謀，進事稱心如意，文武皆能。女命主一生福壽無虧之命。沙中土命，屬相為龍：上天之龍。",
            ["甲戌"] = "守門之犬。口快舌尖，用心不用力，有權柄，名聲四海；夫婦相合；女命旺夫生財，興家立業。35亥山頭火命，屬相為豬：過往之豬。與人和順，初年多災。父母有刑傷，再拜無妨，夫婦無刑皆老，存心中正，豐足興旺，子息有克破，晚兒方好。",
            ["甲申"] = "過樹之猴。一生衣祿不少，性情溫柔。出入皆壓人，初運不利，末運卻佳。夫婦美滿，兒女遲見為好；女命持家，豐物興旺，賢孝之婦45酉泉中水命，屬相為雞：唱午之雞。",
            ["甲辰"] = "為清溫之龍，喜氣春風，出入壓群。利官，近貴，骨肉刑傷，兒女無力；女命主溫良賢達，心直口快，必嫁好夫。大林木命，屬相為蛇：福氣之蛇，有功名之徵兆，伶俐聰明，夫婦命和，做事能遂心願，財豐田足；女命衣良無缺，良婦之命。",
            ["癸丑"] = "海中之牛，少年運氣不好，有父母重婚之兆。夫婦皆好，兒女不孤，愛好九流藝術。爐中火命，屬相為虎：為山林之虎。其性格不定，心直口快，手足忙碌，不利官。女主賢良，聰明伶俐之命。",
            ["癸亥"] = "林下之豬。性剛直不順，不易受欺之人；女命能持家興業，家財豐厚，晚年福壽全美之命。",
            ["癸卯"] = "樹木之兔。酒食不欠，福壽有餘，逢凶化吉，一生平安、興隆一世；女命中末運氣欠佳，操持家務一生。佛燈火命，屬相為龍：有道之龍。待人有禮，安守本分，眾人敬重，一生清閒如意，雖勞不碌，財旺自來；女命賢巧興家。65巳佛燈火命，屬相為蛇：出穴之蛇。",
            ["癸巳"] = "草中之蛇。聰明伶俐，財來不聚，晚年不同前，發財有福；女命賢良，興家助夫。沙中金命，屬相為馬：春中之馬。對人和氣如春風，好交朋友，利官近貴。遇凶化吉，一生無刑。女命能言善語，口尖舌快。55未沙中金命，屬相為羊：敬重之羊。",
            ["癸未"] = "群內之羊。一生有快樂之心，廣行方便，只是救人無恩，反招是非，命中財來不聚，女命賢德持家，晚年榮昌。泉中水命，屬相為猴：過樹之猴。一生衣祿不少，性情溫柔。出入皆壓人，初運不利，末運卻佳。",
            ["癸酉"] = "樓宿之雞。說話流利，巧言善辯，辦事公正，無私心，衣良平穩。六親無緣，晚景旺相，女命益夫助夫興家。山頭火命，屬相為狗：守門之犬。口快舌尖，用心不用力，有權柄，名聲四海；夫婦相合；女命旺夫生財，興家立業。35亥山頭火命，屬相為豬：過往之豬。",
            ["辛丑"] = "路途之牛。對人溫和，初年有意外之危。衣食雖有，但進財不存，老來福壽雙至；女命賢良，旺財之命。金箔金命，屬相為虎：過林之虎。心直口快，不暗地生非。夫婦，兒女有克，遲方好。初年財不聚，事事難好，晚景豐足之命63金箔金命，屬相為兔：樹木之兔。",
            ["辛亥"] = "圈內之豬。一生不貪閒事，不惹是非，早年財物不足，到老榮華；女命主體壯，能興家立業桑松木命，屬相為鼠：葉上之鼠。幼年不災，中年大運大好，衣食足用，主有好妻，身忙心苦，喜憂交替，兄弟少力，六親疏淡，凡事自作自為；女命主賢良，持家之命。",
            ["辛卯"] = "蟾窟之兔。口快，有權。利官近貴身閑，心忙，六親少利，自立自成家業；女命持家，興家旺相。長流水命，屬相為龍：行雨之龍。為人辛勤勞碌，終日手足不停，早歲運悔，晚景財旺發福；女命為一生操持家業，助夫興家。53巳長流水命，屬相為蛇：草中之蛇。",
            ["辛巳"] = "冬藏之蛇。為人機謀多變，志氣人，一生衣食無虧，主有貴人扶持之徵兆，晚年快樂之命。楊柳木命，屬相為馬：軍中之馬。為人勤儉，父母有刑傷，災危可折，早年財物有緊，晚年旺相有餘；女命，主興家也，為賢良之婦。43未楊柳木命，屬相為羊：群內之羊。",
            ["辛未"] = "得祿之羊，一生情性寬大志高，少年不利，夫婦和諧，頭胎見子有刑，女命持家興旺。劍鋒金命，屬相為猴：清秀之猴，情高手巧，聰明伶俐，有智謀，遇事多變。和氣春風，功名有分，有賢德之妻，女主貌美色姿。33酉劍鋒金命，屬相為雞：樓宿之雞。",
            ["辛酉"] = "籠藏之雞。一生伶俐，精神爽朗，口舌能辯，六親生疏，女命主賢德治家，老來榮華之命。大海水命，屬相為狗：雇家之犬。一生好行善，初年平平，末年興旺。",
        };

        // Ch.5 年生肖×月支（逐月考）
        var nayinCh5 = new Dictionary<string, string> {
            ["丑_丑"] = "牛人生於二月，驚蟄之時，此人聞驚迫急，心猿意馬，時欲出頭肯幹，不得其權，自我心強，一生膽大有射虎之威，無人能阻。",
            ["丑_亥"] = "牛人生於臘月，小寒之時，每食生疏而不熟，功名恍惚以難成，有李廣之威力，終是無發安寢樂室，應子傳孫，自身寒薄，通俗一世。",
            ["丑_午"] = "牛人生於七月，立秋之時，豐隆命運高，動力自旺，龍子清高，衣食必新，一生少困難，凡事大吉，四路亨通，利達三江，早有福星照命，晚有孝子孝孫。",
            ["丑_卯"] = "牛人生於四月，小滿之時忙忙碌碌，東走西奔，勞苦無休息，名不得時。",
            ["丑_子"] = "牛人生於正月，新春之時，待時出力，雖目前運滯，後日當有福，自有大用。",
            ["丑_寅"] = "牛人生於三月，清明之時，一翻新氣象，自由自在，到處可居無難，雖有些險岩風波，得能彼岸，他日自有盈餘，獨立心強。",
            ["丑_巳"] = "牛人生於六月，小暑之時，知識才能，令聞令望，半世幸福，半世困苦，東來紫氣運時必到，事業如意，吉多凶少，往來有助，走自由之大路，天賦之美德，吉慶終日，旭日東昇。",
            ["丑_戌"] = "牛人生於冬月，大雪之是地，身寒意冷，有歹命運，待之人用，無法用人，一生災無難蓋乎哉。",
            ["丑_未"] = "牛人生於八月，白露之時，自享祖福，工作時有困難謀事在人，名利平衡，四海流通，有成功之日，進退之日，進退自如，威望隆重，才能發達，技藝精通，健康而身厚重也。",
            ["丑_申"] = "牛人生於九月，寒露之時，萬事如意，智謀權力必集，衣食住行，自有一新，就是出身微賤，有志竟成，謀事如意，長髮其祥，勇往直時，時運平和，排除萬難，一生幸福。",
            ["丑_辰"] = "牛人生於五月，芒種之時，不威不重，破壞艱難重重，慎重為之。",
            ["丑_酉"] = "牛人生於十月，立冬之時，雖身風霜，無衣足食，思想計畫，獨裁自如，事業進步，金錢充裕，負芨有師，有旭日東昇之勢，如月之恒矣。",
            ["亥_丑"] = "豬人生於二月驚蟄之時，能成大功，人傑地靈，再興再望，家聲克振，定家門第，秉性聰慧，有天賦之美德，精力充沛，吉慶終世也。",
            ["亥_亥"] = "豬人生於臘月小寒之時，雖有天賜食祿，奈何無權力，有受天難，凶多吉少，壽元欠長，空有才能。",
            ["亥_午"] = "豬人生於七月立秋之時，衣食豐盈，一生不力自事，不遷就他之事，自專獨裁，凡事一帆風順，勢力強大，吉祥通臨，定然成功。",
            ["亥_卯"] = "豬人生於四月立夏之時，秉性聰敏，職位權貴高尚，雖出有章，但亦有此微困，立足實地，可成大志，無上吉祥。",
            ["亥_子"] = "豬人生於正月新春之時，長生在望，聰敏至賢，威尊望重，利路亨通，有自然之幸福，名利雙收，能成大事業。",
            ["亥_寅"] = "豬人生於三月清明之時，身體最健康，性格剛強，以致與人不和。",
            ["亥_巳"] = "豬人生於六月小暑之時，才高八斗學富五車，怠情成性，缺乏毅力，凡事優柔寡斷，做事無成，無進取之精神，禍福無常，消沉不定。",
            ["亥_戌"] = "豬人生於冬月大雪之時，面貌舒展，易怒，頑固不化，但自尊自福，名利雙全，一世無愁。",
            ["亥_未"] = "豬人生於八月白露之時，自成權威，能為領袖，功業成就，受人敬仰，安樂自尊，但須苦難漸進。",
            ["亥_申"] = "豬人生於九月寒露之時食必豐厚，天賜之祿，受人稱讚，悠悠不斷，逍遙自樂，長髮其祥，忍事耐性，致得榮華富貴。",
            ["亥_辰"] = "豬人生於五月芒種之時，性質溫良，有些小才能，奈何無權力，對藝術文學方面，可發展成功，欲眾大事，無膽略才謀，一生保守，晚景漸佳。",
            ["亥_酉"] = "豬人生於十月立冬之時，榮養漸加，安享其良，人品端厚，受人稱讚，天賦之幸福，富貴成功，晚景昌盛。",
            ["午_丑"] = "馬人生於二月，驚蟄之時，為人清楚乾淨，衣冠必新，但難免狂風雨之災害，時防驚號，有存正德，驚無大害，秉性陪敏過人，四處皆通，自由自在。",
            ["午_亥"] = "馬人生於臘月，小寒之時，一生煩事特多，不得不奔走跋山涉水之苦。",
            ["午_午"] = "馬人生於七月，立秋之時，不冷不熱，無凶無禍，靈感過人，聰敏至極，溫柔文雅，技巧藝術超人，意志常欲成功，有日必得貴人助，誘人之力特強，精力充沛，萬事如意，一生無大災害。",
            ["午_卯"] = "馬人生於四月，立夏之時，必有赴湯蹈火之焚，負擔很重，走東到西，錢來錢去，忙忙碌碌，苦惱一世，不得人憐，憂愁不絕，無成功之日。",
            ["午_子"] = "馬人生於正月，新春之時，精神爽快，活龍活虎，浩浩蕩蕩、欣欣向榮，但雪地將開，冰天將放，一生少得嘗新之機。",
            ["午_寅"] = "馬人生於三月，清明之時，雲開雨散，勢有千里之途，耀武揚威，文沖鬥牛，志氣淩雲，無困難災害，英豪人傑，膽識過人，令人敬仰，鄰里稱美，遠悅近和，謀事大成，凡試必先登。",
            ["午_巳"] = "馬人生於六月，小暑之時，粗茶淡飯，終身疲勞，意外過慮。",
            ["午_戌"] = "馬人生於冬月，大雪之時，冰天雪地，寸步難行，負理累累，奔波勞碌，一世遇小安寧之日也，凡事多難。",
            ["午_未"] = "馬人生於八月，白露之時，技藝智謀過人，文武兼全，膽大心虛，德高義衙，感情和合，有志氣，勇往直前，凡事逐漸成功，腳踩樓梯，無上吉祥，是非兼半，紫氣東來，俊氣才能，英傑才人。",
            ["午_申"] = "馬人生於九月，寒露之時，與人同樂，才能智力均尚，意志不堅，事不遂心，自然自安，不願處世學好，一遇挫折，灰心不起。",
            ["午_辰"] = "馬人生於五月，芒種之時，門遷新氣象，堂宇舊規模，創基立業，千辛萬苦，家道興隆，福祿得長生。",
            ["午_酉"] = "馬人生於十月，立冬之時，起初災害並至，而晚景是幸福偕來，但變動異常，排除萬難，成功者有之，而挫折者有之，此為歷來之英傑人物，秉性穎悟，俠膽義心，臨事難達目的，一生平凡之象。",
            ["卯_丑"] = "兔人生於二月，驚蟄之時，興高采烈，光彰裕後，前途有望。",
            ["卯_亥"] = "兔人生於臘月，小寒之時，年終歲畢，衣祿輕淡。",
            ["卯_午"] = "兔人生於七月，立秋之時，巧期佳節，男女聰敏顯耀，逍遙自在，清高顯耀，青雲之志，白屋之人，不難求望，諸凡遂心。",
            ["卯_卯"] = "兔人生於四月，立夏之時，足衣足食，文質彬彬，威望隆重，大將之材，名利兼旺，夫妻榮華，子孫顯貴，一帆風順，技藝精通，快樂逍遙，精神爽爽，不費心神，所謀如欲，富貴吉祥。",
            ["卯_子"] = "兔人生於正月，新春之時，奔波用功，忙忙碌碌，終日若得，尚稱命運高。",
            ["卯_寅"] = "兔人生於三月，清明之時，躍武揚威，聰敏活潑，有沖天之勢，儼然新氣象，欣欣向榮，如月之恒，如日之升，智略權謀，勢如破竹，時運必到，凡謀必遂，能成大業，一生幸福，長髮其祥，大啟爾宇。",
            ["卯_巳"] = "兔人生於六月，小暑之時，萬象更新，威尊名望，利呼亨通，健康富貴，事業順時，有專主這才，榮華子息，賢妻能夫，人傑地靈，福壽康寧。",
            ["卯_戌"] = "兔人生於冬月，大雪之時，遍地皆白，有志無機，苦悶煩惱，險災勞困，一世順利少有，困難辛苦。",
            ["卯_未"] = "兔人生於八月，白露之時，心柔性和，善養德行，無欺無詐，秉性聰敏，順時逆退之靈要應變力特強，能撫助眾人，忠心挾放，處世有方。",
            ["卯_申"] = "兔人生於九月，寒露之時，事業順時，時感喪心晦氣怠惰不力，不爭取，不立志，終身幸耳。",
            ["卯_辰"] = "兔人生於五月，芒種之時，事業順利發達，經營風順，受人重視，有宇宙之大精神，互助合作，處世有情，待人接物，恭恭有禮大吉大利。",
            ["卯_酉"] = "兔人生於十月，立冬之時，青青地中禾，不力自食足，無須奔波，靜居養性，永保萬事，雖有此驚險，忍之無事，不願處世交友，不謀不就，清高一生，順利成矣。",
            ["子_丑"] = "鼠人生於二月，驚蟄之時，此人膽大怕驚，終是小膽，一生令人可愛，文雅溫和，可得多貴人，處世大方，雖不能掛師掛將，文印可保，腦海聰敏，瞭若指掌，在千辛萬苦，大驚大赫之下樂登彼岸。",
            ["子_亥"] = "鼠人生於臘月，小寒之時，歲在年末，五穀糙糧都作熟糧，儲酒淹肉，以待除夕，應是吃好，安然康泰，飽暖一生。只是地凍天寒，辛勤勞作方有福。",
            ["子_午"] = "鼠人生於七月，立秋之時，五穀豐登，需防小人陷害，不可糊塗。",
            ["子_卯"] = "鼠人生於四月，立夏之時，一生衣祿輕淡，勞碌奔波，不得貴人，處處遇難，不時小人搗亂，就驚險重重。",
            ["子_子"] = "鼠人生於正月，新春之時，萬戶窮富，家家均是葷。當食之間，豐衣足食，一生口神速不輕。每食必葷，但食太多，往往心生糊塗。",
            ["子_寅"] = "鼠人生於三月，清明之時，一生覺見喪門，到處碰枯骨，雖有酒肉之食祿，無奈總是流淚痛心之事。大則不平之事，能致死於非命，小則不平之事，可能出家，次之必成庸人。",
            ["子_巳"] = "鼠人生於六月，小暑之時，熱氣陽陽，光明樂道，遍地皆春，但防烈日逼不，身遭其疾，雖是乾渴，宜居於江河湖海之畔，多尋水源，一生出頭人上。",
            ["子_戌"] = "鼠人生於冬月，大雪之時，必得登家，白吃山崩，待時而動，糧食賴祖。",
            ["子_未"] = "鼠人生於八月，白露之時，不但衣食盈餘，並得多重美賢之夫妻，相貌出眾，舉筆成章，定作尚客，並常赴約喜筵，百事可成，多有貴人扶持，終有成功之日，切莫自妥，必成大器也。",
            ["子_申"] = "鼠人生於九月，寒露之時，為人穩重，怕出頭做事，畏寒畏冷，可能衣食保暖，決難成大將之才。總得美賢之夫妻，難免口舌。",
            ["子_辰"] = "鼠人生於五月，芒種之時，東走西奔，勤能忘食，不失良機，能成富室，一生不怕淒涼，不畏艱難可獨立成家，能成萬眾崇拜，可作人之模範，老來自福，始終無難。",
            ["子_酉"] = "鼠人生於十月，立冬之時，五穀歸倉，處處糧空，四處尋吃，多有依人為生，雖有沖天之志，總少機會。",
            ["寅_丑"] = "虎人生於二月驚蟄之時，正是出力之時，用武這地，有掀天揭地這奇才，建功業，奏奇功，能統率眾人，智略權謀，勢如破竹。",
            ["寅_亥"] = "虎人生於臘月小寒之時憂悶頻多，凶多吉少，多遇暗箭，避之不及，衣食不周，時防不意殺身之禍，步上慘苦這境，萬事挫折，行動不便，腚安分守己，慎之祥也。",
            ["寅_午"] = "虎人生於七月立秋之時，秋天老虎，格外厲害，脾氣剛強，意志堅強，跋山涉水之苦，如折枝之易，赴湯蹈火之難，在所有辭，能克服萬難，凡事可成矣。",
            ["寅_卯"] = "虎人生於四月立夏之時，清和風暖，雪游四海，門庭熱鬧，家道興隆，出將入相之為，家庭和睦，骨肉相親，凡桂五枝勞，眥荊林茂，育子皆賢，養氣成貴，一生無煩惱，處處亨通矣。",
            ["寅_子"] = "虎人生於正月新春之時，萬象更新，虎躍虎威，名揚四海，吼聲振天，家運隆昌，富貴吉祥，子孝孫賢。",
            ["寅_寅"] = "虎人生於三月清明之時，更是出頭之天，自成權威，受人敬仰，沖天之勢，立大功勞，青雲有路，當際風雲之會，必承雨露之恩，門庭新氣象，堂宇換規模。",
            ["寅_巳"] = "虎人生於六月小暑飲水無源，到處有難，氣沖鬥牛，呱呱亂叫，處處少人，不足氣力有餘，受天之能難，所謀不遂，每多因難。",
            ["寅_戌"] = "虎人生於冬月大雪之時，四方皆敵，出行艱難，憂愁不絕命途外舛，凡事小心可也。",
            ["寅_未"] = "虎人生於八月白露之時，亦正榮耀之時，先知先覺之才能，貫徹始終，利達四海，一本撐住天下，為眾人敬，威鎮人群，萬事如意，天賜之福，聰敏活潑，文章蓋世，不屈不撓，足立實地，名揚四海，聲振山河。",
            ["寅_申"] = "虎人生於九月寒露之時，漸沽其勢也，自立心欠強，依人穿吃，處處不通，所謀不遂，喪心悶氣，有力不出，坐吃山空，工作鬆懈，絕俗離塵，野外之人，總清高亦是弱懦，負重不得，擔草不起，滿腹經綸，生不逢辰，只思守，不想鴻圖大展。",
            ["寅_辰"] = "虎人生於五月芒種之時，有機可為，凡謀皆就，自食其力，有尊嚴風度，進退能自由，文武兼能，膽略過人，建立基業，雅量敦直，性格剛強，以人不和，勇往直前，竟到成功之地，權力勢焰，受人為難，白手成家，富貴成功。",
            ["寅_酉"] = "虎人生於十月立冬之時，祿馬分鄉，勞碌奔波，求謀多戾創業維艱，獨力難持，因人創立，秉性聰敏，義氣溫和，可做可維，順，時聽命，成者自成也，無道培之功，舉器不凡庸，奈非時也。",
            ["巳_丑"] = "蛇人生於二月，驚蟄之時，眠中驚醒，混混噩噩，不知生向，性情不高，志氣萎頓，作事懶成，獨立難持，總請高，亦是寒懦，祿薄福輕，煩事纏綿。",
            ["巳_亥"] = "蛇人生於臘月，小寒之時，歲寒冰迸發，雪花六出，絕俗離塵，修德有功，煩悶苦悸，難享幸福。",
            ["巳_午"] = "蛇人生於七月，立秋之時，逍遙自在，立業可期，凡事如願，智勇雙全，敏捷聰明過人，經營有道，財恒足矣，性情溫柔，道德高尚，功業成就，受人敬仰。",
            ["巳_卯"] = "蛇人生於四月，立夏之時，精氣浩大，威震四方，雅量敦厚，秉性聰明，地位權貴至高，有大志大能，俠義心腸，眾人敬仰，受天之福。",
            ["巳_子"] = "蛇人生於正月，新春之時，陽氣將盛，活躍起來，四出有路，但經風霜之苦，難免事事如麻，每每欲一步登天，時機未到。",
            ["巳_寅"] = "蛇人生於三月，清明之時，聰敏巧能，可圖僥倖，青雲之志，白屋之人，一舉成名，為國之賢，功在四方，利達三江，能成大事大業，多勞多功，精神爽快，謀事諸遂。",
            ["巳_巳"] = "蛇人生於六月，小暑之時，萬事如意有天賜之福，達成功之日，合謀共人，互助有力，夫妻相榮，德量才能兼備，意志堅銳，熱誠忠厚，慈祥有德，吉祥至尚，名利皆就。",
            ["巳_戌"] = "蛇人生於冬月，大雪之時，八方晴光，四處悉雪白，進出無路，清閒淡薄，含辛茹苦。",
            ["巳_未"] = "蛇人生於八月，白露之時，忠厚傳家，和鄰睦戚，愛親敬長，兄友親恭，鄉里稱道，美德美善，忍事柔性，建立基業，自然之幸福。",
            ["巳_申"] = "蛇人生於九月，寒露之時，一倍工夫一倍熟，及時耕各及時收，以和為貴，惟靜惟佳，曾經落地之關，自有登天之日，秋霜意氣和。",
            ["巳_辰"] = "蛇人生於五月，芒種之時，膽力才謀過人，能克服萬難功利榮在，能察時世，有先智先能，哲人之頭腦，一生平安，貴格之造，福祿悠久，家運隆昌。",
            ["巳_酉"] = "蛇人生於十月，立冬之時，無妻賢明，雖有不俗之志，決有忠心之感，樂於助人，善於人交，家空財薄，待修有期，而陰霧迷空。",
            ["戌_丑"] = "狗人生於二月，一旦定下目標之後，卻能接受而努力不懈，個性外向、樂觀且開朗，人緣佳而受人愛戴，狗年出生的人基本上都具有敏銳的直覺力和判斷力，是邁向成功的一大武器。但是，凡事過於樂觀，就難免會輕敵而種下失敗之因。對人信任，做事缺乏周密的計畫，往往會中途就宣告失敗而白忙一場。",
            ["戌_亥"] = "狗人生於臘月，感情豐富，容易博得別人同情而幫助他，愛情方面，一旦發現目標，就窮追不捨，容易陷入熱戀中完全失去理智。男性略帶神經質，時常擔心別人對自己的看法如何，自尊心極強，是一個不服輸凡事喜歡與人競爭到底的人。對人和善，絕不會做出出賣朋友的事，所以瞭解他的人，能與之結為好友。",
            ["戌_午"] = "狗人生於七月，處理事情幾乎無懈可擊，所以不管走到哪裡都會受人重用。女性多難保守秘密，男性則具有開朗而善良的個性，多情善感，往往會引出綺麗的戀情，但是缺乏耐心，常常不了了之。晚婚的情形較多。緣份方面，該月男性以辰年三月出生的女性為妻最為適合。",
            ["戌_卯"] = "狗人生於四月，喜歡講理論，好思索，時常以思想家自居，野心大，充滿熱情而又有積極的奮鬥力。這個月出生的人，就像一個快速轉動的馬達，戰鬥力高昂。自信高，對於自己所講的話，覺得很有道理而會頻頻點頭。參加朋友婚宴，會與同坐一桌的人聊得很開心，進而結為朋友，甚至能從中找到自己未來的人生伴侶。",
            ["戌_子"] = "狗人生於正月，受到上一年的影響，具有悠然的心性且重感情。開始運勢不好時，需要默默努力，一步一步向前邁進，才能開創自己的錦繡前程。適合的職業有：教職員，編輯，技術人員等。這個月出生的人，有節約儲蓄的美德。無論男女，對愛情都抱著慎重的態度，喜歡細水和的感情，而不會沉迷于一時激情。",
            ["戌_寅"] = "狗人生於三月，精神旺盛，能將理想和現實劃分清楚。在文化、藝術方面有特長，喜歡幫助朋友。適合的行業：設計師或服務性質的工作。女性最適合的物件為有積極行動力的男性。子年八月或未年十月的就較適合。",
            ["戌_巳"] = "狗人生於六月，起步比別人早且順利，年輕時就得到上司和長輩的提拔，擔任重要的職務，聰明伶俐、敏捷，膽識過人，只可自視過高，而不懂謙讓的美德。喜歡旅行，也喜歡過著華麗的生活，因而開支大，朋友雖然很多，但多半交情不深，不喜歡向人借錢，手邊有多少就用多少，所以也沒有什麼積蓄。",
            ["戌_戌"] = "狗人生於冬月，對工作兼具鬥志和沉著，是個很能幹的人，對人觀察很敏銳，周圍的朋友有一點小差異，立刻能發現。但是過度觀察，在戀愛時就變成了狂疑、忌妒、愛他（她）卻不信任對方，不僅自己苦惱也會讓對方受不了。對於婚姻猶豫不決，因此常錯過良緣。適合的職業：秘書、新聞記者等。",
            ["戌_未"] = "狗人生於八月，具有冷靜和喜歡批評人的性格，因此其思想總是比行動走在前面，天性不好動，適合從事於內勤工作。做事情考慮過多，缺乏魄力和勇氣，所以往往在最後關頭敗給了對手。一旦遇到挫折就心灰意冷，缺乏耐心和繼續前進的鬥志。雖然缺點不少，但是體貼細心都是這個月出生的人最大的優點。",
            ["戌_申"] = "狗人生於九月，這個月出生的人思想開放，喜歡嘗試新事物，故在私生活方面較為隨便，對於個人喜惡十分極端，不善於言辭，也不慣於壓抑感情，喜怒哀樂完全表露有臉上，所以容易得罪朋友。這個月出生的人，對於酒精情有獨鍾，十分好飲但能自我節制。記憶力特別好，好比是一個沒打過的電話號碼，他仍記得很清楚。",
            ["戌_辰"] = "狗人生於五月，頭腦靈活與可親的態度，是這個月出生的人的最大優點，雖然對人坦誠，但卻頗會拍馬屁或利用關係爬上高位。時常面帶微笑，諸如\"請多提拔\"之類的話經常掛在嘴邊。所以不是領導人物，而適合做秘書或參謀等工作，年輕時運勢較差，凡事不可過於強求，一旦過了三十歲，就能逐漸出人頭地。",
            ["戌_酉"] = "狗人生於十月，為人正直、親切、聰明伶俐，善於交際，能識別利害關係，屬於\"行動派\"的人，一生的運勢不錯，年輕時能嶄露頭角。女性人緣頗佳，無論和什麼人都相處融洽。有時反而被人認為是多管閒事，平時生活簡樸，不太重視物質的享受。儘管如此，在她身上卻往往流露\"貴夫人\"或\"名門淑媛\"的氣質。",
            ["未_丑"] = "羊人生於二月驚蟄之時，秉性溫和，處理有方，衣食無憂，四路皆通，萬事多利。",
            ["未_亥"] = "羊人生於臘月小寒之時，寸步難行，前途渺茫，災難多多，災禍重重，時受有嚴守，行動無自由，終生奮鬥，晚年享福。",
            ["未_午"] = "羊人生於七月主秋之時智勇雙全，意志堅銳時運必至，萬事如意，能成大事業。",
            ["未_卯"] = "羊人生於四月立夏之時，災害常有，晚福必來。",
            ["未_子"] = "羊人生於正月新春之時，三陽開泰，喜氣盈門；詞訟是非能免，天災地變常遇，晦氣多多。",
            ["未_寅"] = "羊人生於三月清明之時，聰敏出眾，財源利達，萬事吉祥，衣食必俠，富貴健康，享自然之幸福。",
            ["未_巳"] = "羊人生於六月小暑之時，熱心忠直，受天之福，名種榮達，進退無難，處世和平，合謀志同道合，言而有信；家運隆昌，榮華一世，晚享子福。",
            ["未_戌"] = "羊人生於冬月大雪之時，謀事難成，財邊艱難，百事不遂，千頭萬結，中年始興。",
            ["未_未"] = "羊人生於八月白露之時，漸入佳境，爵位升遷，然誹諺之多招，災之不免，但有天之保，受主之力自行獨正，自為有道，心正不怕邪。",
            ["未_申"] = "羊人生於九月寒露之外，作事有成，但根基淺弱，創業艱難，小人常寒，病疾不斷，辛勞奔波，如以穩中求進，也是平安一生。",
            ["未_辰"] = "羊人生於五月芒種之時，財豐利足，權勢必高，所謀如意，出言有音，意志堅定，腳踏實地；能統率眾人，受人之難，亦受敬仰。",
            ["未_酉"] = "羊人生於十月主冬之時，小陽春暖，應是了頭之日。",
            ["申_丑"] = "猴人生於二月，驚蟄之時，多憂驚怪，難化愚頑，雖有雨露之恩，官居不久，而事宜囊篋空虛。",
            ["申_亥"] = "猴人生於臘月，小寒之時，大雪封北京時間，冷冷冰冰，動搖不安，無謀淺慮，慘澹經營晚景平安。",
            ["申_午"] = "猴人生於七月，立秋之時，五穀豐登，一倍工夫一倍熱，及是耕種及時收，安然自在度平生。",
            ["申_卯"] = "猴人生於四月，立夏之時，忙忙碌碌，辛勤勞苦，知識才能，遠近聞名，一生榮華，事從心欲，吉凶兼半，互助有人，身體又鑠，衣食飽暖。",
            ["申_子"] = "猴人生於正月，新春之時，活躍力甚佳，精神爽快，事業出外，浮沉小病難免。",
            ["申_寅"] = "猴人生於三月，清明之時，才能智技過人，江湖有伴，鄉鄰有田，婚姻美滿，終身是幸，稼穡難成，經營可企，凡事順來逆去，改禍呈祥，前途光明。",
            ["申_巳"] = "猴人生於六月，小暑之時，清和日炎，風調氣溫，身居得到，爽氣滿身，樂樂無憂，唯粗衣淡食不缺，解菜美肴難得，吉凶禍福平均，婚姻自由美滿，子孫賢貴兼能。",
            ["申_戌"] = "猴人生於冬月，大雪之時，憂悶多起，時有凶兆，精神萎頓，性情乘忤。",
            ["申_未"] = "猴人生於八月，白露之時，事業享通，能成大事業，前途榮進，福祿無窮，不能富貴亦賢望，天賦美德，家門和合。",
            ["申_申"] = "猴人生於九月，寒露之時，才能功成，自強心高，但半途受挫而致失敗，是非兼半，吉凶並行，時能發展成功，事與願違，一生平凡。",
            ["申_辰"] = "猴人生於五月，芒種之時，東奔西跑，赴湯蹈火，奮鬥有企，衣祿輕淡，待人接物，恭敬有禮，進退自如，諸謀可成，是非口舌小有，一生承享自力自食，自求自謀也。",
            ["申_酉"] = "猴人生於十月，平凡成家，若不安份，不得如願。",
            ["辰_丑"] = "龍人生於二月，又名\"杏月\"，時值\"仲春\"。",
            ["辰_亥"] = "龍人生於臘月，又名\"臘月\"、\"除月\"，時值\"季冬\"。\"龍\"年十二月出生的人，成功之地，一生得志，辛勞也。",
            ["辰_午"] = "龍人生於七月，又名\"蘭月\"、\"瓜月\"，時值\"孟秋\"。\"龍\"年七月出生的人，事事成就，福祿悠久。",
            ["辰_卯"] = "龍人生於四月，又名\"槐月\"，時值\"孟夏\"。",
            ["辰_子"] = "龍人生於正月，又名\"正月\"、\"元月\"，時值\"孟春\"。",
            ["辰_寅"] = "龍人生於三月，又名\"桃月\"、\"蠶月\"，時值\"季春\"。",
            ["辰_巳"] = "龍人生於六月，又名\"荷月\"、\"暑月\"，時值\"季夏\"。",
            ["辰_戌"] = "龍人生於冬月，又名\"冬月\"，時值\"仲冬\"。",
            ["辰_未"] = "龍人生於八月，又名\"桂月\"，時值\"仲秋\"。",
            ["辰_申"] = "龍人生於九月，又名\"菊月\"，時值\"季秋\"。\"龍\"年九月出生的人，謀事如願，雅量敦厚，偶有躁性，固執剛強，好除暴安良，勇往直前，不枉己志。龍人生幹十月，又名\"小春月\"，時值\"孟冬\"。",
            ["辰_辰"] = "龍人生於五月，又名\"榴月\"，時值\"仲夏\"。\"龍\"年五月出生的人，正活動之時，權力勢焰，利路亨通。",
            ["酉_丑"] = "雞人生於二月，是個愛好自由，十分厭惡束縛的人，凡事不肯讓步，喜歡力爭到底。男性是很專情的人，對自己喜歡的女性用情很深。女性於自己的穿著打扮，頗為注重，但是不會花太多的金錢或時間在上面。此月出生的男性適合的對象為子年八月出生的女性。女性則為辰年六月或亥年九月出生的人。",
            ["酉_亥"] = "雞人生於臘月，是一個大方、開朗、而且有手腕的人，善於攻心，說服他人的能力很強，喜歡幫助他人，即使是無理的要求，他多半也會答應。女性喜歡刺激的生活，宜於亥年八月出生的人結合，男性則應選擇戌年十月出生的為其伴侶。適合的職業有：政治家，新聞記者，需要有敏銳的觀察力和判斷力的工作。",
            ["酉_午"] = "雞人生於七月，頭腦靈活，不時有新的點子和構想，且能一一實行。只是性子太急，缺乏冷靜和沉著，所以顯得不夠成熟。適合的職業有貿易進出口，電器事業，出版製作等，帶有投機性質的工作。對於感情十分執著，用情又深，對自己喜愛的人會奉獻出一切，女性適合個性溫和而能包容她的男性，丑年十二月出生的男性最為符合。",
            ["酉_卯"] = "雞人生於四月，在所有雞年出生的人之中，做事從容不迫，且生性樂觀。頭腦聰明而有自製心，有浪費金錢的習慣，雖然如此，仍頗有積蓄這個月出生的人，*白手起家而成巨富的例子很多。適合使用頭腦的職業，尤其是金錢出入多的事業。喜歡幫助別人，也善於照顧別人，天生樂觀。女性做事腳踏實地，未年十二月出生的是美滿的一對。",
            ["酉_子"] = "雞人生於正月，具有敏銳的直覺，才智也不錯，且精於打算，甚為健談，因此頗受人歡迎。但是，有時喜愛誇大其詞，往往\"語不驚人死不休\"，所以日子久了，別人對他的話通常會打三折，受吹牛，講究派頭，常說些不切實際的話，屬於\"紙上談兵\"的人，喜歡到處國旅行，尤其是歐洲，善於理財，在經濟上頗為充裕，生活舒適。",
            ["酉_寅"] = "雞人生於三月，是個理想主義者，富有正義感，但是略帶神經質，對於金錢不太重視，且有\"浪費之癖\"，多半沒有積蓄。這個月出生的人官運不錯，短短幾年即可爬到高位，自己開店做生意也很合適，但凡事得*自己，較為辛苦。無論男女對感情都十分執著，一旦得不到對方，可能因愛生恨而釀成悲劇。",
            ["酉_巳"] = "雞人生於六月，感情十分豐富又敏感，對於細小的事也會哭泣或發怒，也是一個喜歡計畫的人，但美中不足的是缺乏耐心，常常前功盡棄。適合的職業有：演藝人員，製作人都可。緣份方面，女性宜與寅年九月或子年十月男子為伴，男笥宜與丑年八月女性結合。",
            ["酉_戌"] = "雞人生於冬月，不管年紀多朋，仍能保有一顆赤子之心，天真純潔而沒有心機，對任何事都能充滿著好奇心。男性對愛情的態度，抱著可有可無的態度。屬?quot;獨身主義\"者。但是，一旦結了婚，態度就不再那麼冷淡，將是個幽默又體貼的好丈夫。",
            ["酉_未"] = "雞人生於八月，為人大方慷慨，是自由放任主義者，對於別人的過錯，有一顆寬容的心，同樣地他也很會為自己找藉口，事事都不苛求，意志薄弱，容易受外界誘惑。有藝術方面的天份。女性有事業心，多半是職業婦女。個性浪漫，又有點孩子氣，有良好的鑒賞力，適合與寅年九月出生的男性結合。",
            ["酉_申"] = "雞人生於九月，想法奇特，討厭模仿他人，所以一舉一動都喜歡標新立異。這個月出生的男性，頗有才幹，所以選擇\"唯才是用\"的公司，發展較大。女性是個性隨和，不會裝模作樣。外型清麗，所以很有魅力。男性最好選擇純樸來自鄉下的女孩為妻，都市里長大的時髦女較不合適。丑年四月或午年十一月出生的人，都是理想的對象。",
            ["酉_辰"] = "雞人生於五月，是個勤勉而溫和的人，但是，凡事過於小心，做事缺乏魄力，以至一無所成。朋友較少，生活的圈子也較窄，為人誠實，是最佳的總務人選。適合的職業有：圖書管理員，政治家，行政秘書等，不論從事何等工作，都不宜留在鄉里，離鄉背景出外打工比較有發展。愛情方面順利，不會有太大的波折。",
            ["酉_酉"] = "雞人生於十月，認真而溫厚，內心意志堅強，凡事不服輸，是個外柔內剛的人。無論做什麼事，都會按部就班，一步步向目標邁進，所以最後的勝利往往屬於他。女性頗得從緣，但是較缺乏容人的雅量，有時為了一點小理就和人家吵不休，或將對方批評得一無是處。相信命運，喜歡四處看手相。",
        };

        // Ch.6 年生肖×日支（逐日考）
        var nayinCh6 = new Dictionary<string, string> {
            ["丑_丑"] = "牛人生於丑日，鵬程萬里，須明症侯，思神過慮，不免心疼眼花。",
            ["丑_亥"] = "牛人生於亥日，終日功在沙場，日日作吊人，命中雖強，宜早歸寧。",
            ["丑_午"] = "牛人生於午日，幸有月德照臨。",
            ["丑_卯"] = "牛人生於卯日，小忿小懲必到，爭長競短大虧，一生小幸。",
            ["丑_子"] = "牛人生於子日，雖有些小微病，終是不藥而痊。",
            ["丑_寅"] = "牛人生於寅日，太陽高照，五心不開，時有煩惱，空中懸人之象，每思萬空之古感，男有登科遇美，女有子歸之才。",
            ["丑_巳"] = "牛人生於巳日，財路亨通，交友多指背，當是貴人。",
            ["丑_戌"] = "牛人生於戌日，天德福星常照，躍紀沿途，四方求名利，有日不歸家。",
            ["丑_未"] = "牛人生於未日，雖有貫之根基，每事不詳無善德後果，一世多難，宜守之。",
            ["丑_申"] = "牛人生於申日，紫徽高照，求婚多遂，雖有不意之事，貴人照臨，鼇頭可占，青錢得中，事在人上。",
            ["丑_辰"] = "牛人生於辰日，太陽高照，每有女貴人，多憂榮絆，少得安靜，陽逆陰和之兆。",
            ["丑_酉"] = "牛人生於酉日，命帶將星，雖誹謗多招，災殃不免，虎威旺重，克制鬼賊，成之在矣",
            ["亥_丑"] = "豬人生於丑日，太陽星高照，雖有不祥兆，得貴人改之。",
            ["亥_亥"] = "豬人生於亥日，浮沉不定，事業飄零，多起意外，交友不利。",
            ["亥_午"] = "豬人生於午日，紫徽星高照，龍德臨命，諸事迪吉，求謀多利，時有天難，逢凶化吉。",
            ["亥_卯"] = "豬人生於卯日，事業高尚，財利如山，一生富裕，但小人所陷，時有官非。",
            ["亥_子"] = "豬人生於子日，太陽星高照，凡事逢凶化吉，命坐風流有感晦氣悶心。",
            ["亥_寅"] = "豬人生於寅日，一生是非，絲絲不斷，麻煩臨身難脫。",
            ["亥_巳"] = "豬人生於巳日，驛馬坐命，朝東暮西，當有一發之時，破耗難免。",
            ["亥_戌"] = "豬人生於戌日，單身只影居時多，病患常有。",
            ["亥_未"] = "豬人生於未日，命帶白虎，破財幾重，但技藝才能過人，溫雅淑慧，聰敏出眾。",
            ["亥_申"] = "豬人生於申日，口舌較多，生意無聊，福星天德二星保命，終身平坦樂道。",
            ["亥_辰"] = "豬人生於辰日，月德臨命，出力為事，喜事重重，偶有小病，須防小破財。",
            ["亥_酉"] = "豬人生於酉日，口舌破碎，令人費解。",
            ["午_丑"] = "午人生於丑日，紫徽星高照，謀事有成，終身少難，取權地位高尚，當有處事困難之時，引來晦氣重重。",
            ["午_亥"] = "午人生於亥日，生意無聊，大敗之時，幸月德來臨。",
            ["午_午"] = "午人生於午日，命帶將星，職掌大權，躍馬沿途，生財有道，常感傷神。",
            ["午_卯"] = "午人生於卯日，天喜星、福星臨命，諸事迪吉，求謀順遂，一帆風順，建業立家速成，偶有是非口舌發生。",
            ["午_子"] = "午人生於子日，災殺歲破必來，是非口舌常有。",
            ["午_寅"] = "午人生於寅日，白虎臨命，交友多犯指背星，諸事不太如願。",
            ["午_巳"] = "午人生於巳日，身體欠康，時有破碎，東倒西歪，坐臥不安，萬里可行，頭疼眼花難免。",
            ["午_戌"] = "午人生於戌日，聰敏坐學堂，文名四海，小人四起。",
            ["午_未"] = "午人生於未日，太陽星高照，必出遠方辦事，時有天南地北，凡事不逆。",
            ["午_申"] = "午人生於申日，命帶驛馬，一發如雷，一氣沖天，如虎下山，時常形單影隻。",
            ["午_辰"] = "午人生於辰日，浮沉不定，漂流事業，孤身雙影，宜少出早歸。",
            ["午_酉"] = "午人生於酉日，太陽星高照，紅鸞喜事多，煩惱事多起。",
            ["卯_丑"] = "卯人生於丑日，單身零丁，凡事多不遂矣。",
            ["卯_亥"] = "卯人生於亥日，白虎臨命，流連少短，不幸中不大成。",
            ["卯_午"] = "卯人生於午日，太陽星高照，得賢妻之助，雖有不測之事，尚無大害矣。",
            ["卯_卯"] = "卯人生於卯日，命帶將星，官帶俸祿，金匱如山。",
            ["卯_子"] = "卯人生於子日，問酒討杏，應見紅鸞之喜，並有添子應孫之吉兆。",
            ["卯_寅"] = "卯人生於寅日，歲犯天官符，是非口舌，一生難免。",
            ["卯_巳"] = "卯人生於巳日，命帶驛馬，主鄉避井，事業尚稱順利，不免小不利。",
            ["卯_戌"] = "卯人生於戌日，紫徽星照命，順多逆少，事在人上。",
            ["卯_未"] = "卯人生於未日，智慧玲瓏，文藝精通，時有小人，造成浮沉不定。",
            ["卯_申"] = "卯人生於申日，身體欠強，幸有月德，雖死有生也。",
            ["卯_辰"] = "卯人生於辰日，太陽星坐命，出事可成，到處皆通。",
            ["卯_酉"] = "卯人生於酉日，月空歲破大耗，常有不幸。",
            ["子_丑"] = "鼠人生於丑日，子與丑合，親子和睦，家和人和，處處可喜。",
            ["子_亥"] = "鼠人生於亥日，多見不樂。牛",
            ["子_午"] = "鼠人生於午日，子午正沖，每遇凶難，不得其利。",
            ["子_卯"] = "鼠人生於卯日，必得祖蔭，並承夫妻之助，一生安享福祿。",
            ["子_子"] = "鼠人生於子日，命帶將星，當有機出眾做事，就是光宗耀祖，爭氣子孫。",
            ["子_寅"] = "鼠人生於寅日，是為驛馬，必須終日奔走四方，難有歸家，多作異鄉之客。",
            ["子_巳"] = "鼠人生於巳日，宜在衣食住行方面，多加注意。",
            ["子_戌"] = "鼠人生於戌日，時欲單身匹馬。",
            ["子_未"] = "鼠人生於未日，未午克制子水，自我心強。",
            ["子_申"] = "鼠人生於申日，申金克制子水，當為人間孝子，思親甚切。",
            ["子_辰"] = "鼠人生於辰日，華蓋坐命，聰明敏捷過人，必能名登金榜，螯頭獨佔。",
            ["子_酉"] = "鼠人生於酉日，命坐桃花，每欲花前月下，時暮歌舞佳人，風流才子。",
            ["寅_丑"] = "寅人生於丑日，紅鸞星高照，瑞靄盈門。",
            ["寅_亥"] = "寅人生於亥日，破落失敗，一馬千里，奔走四方，無可歸也。",
            ["寅_午"] = "寅人生於午日，命帶將星，往來無白丁，時有官鬼小人，搗亂紛紛。",
            ["寅_卯"] = "寅人生於卯日，太陽星高照，新鮮蓬勃，時有空虛之感。",
            ["寅_子"] = "寅人生於子日，命犯天狗星，狼狽不堪。",
            ["寅_寅"] = "寅人生於寅日，命犯太歲，乍冷乍熱，身體不寧也。",
            ["寅_巳"] = "寅人生於巳日，太陽臨門，其欲感太深，時有麻煩之事，孤神相伴。",
            ["寅_戌"] = "寅人生於戌日，福星照命，應享祖基之福，悉是為國增光，為民造福，單身漂零時多。",
            ["寅_未"] = "寅人生於未日，月德照臨，萬事可行，時來天喜星至，不是嫁娶，就是見子見孫，魚門三經浪，躍馬四方，令人敬仰。",
            ["寅_申"] = "寅人生於申日，雖有名旺，困難頗多，百事小心。",
            ["寅_辰"] = "寅人生於辰日，虎頭蛇尾。",
            ["寅_酉"] = "寅人生於酉日，紫徽星高照，東來紫卸，結群行做，百事可成，雖有小破碎，不在其意。",
            ["巳_丑"] = "巳人生於丑日，聰敏賢能，技術高超，一生難免白虎，凡事宜慎之。",
            ["巳_亥"] = "巳人生於亥日，驛日之命，異鄉之客，千里之人，凡事多勞。",
            ["巳_午"] = "巳人生於午日，喜逐多情，時感悶氣，幸有太陽星高照，逢災化祥",
            ["巳_卯"] = "巳人生於卯日，愁悶苦惱，諸事小難，狼狽有時。",
            ["巳_子"] = "巳人生於子日，凡事都吉利，一生喜事重重，雖然不時有天定厄運，但並無大難。",
            ["巳_寅"] = "巳人生於寅日，交友不利，生意可做，財利正旺，謀事大吉，身強力壯。",
            ["巳_巳"] = "巳人生於巳日，交友犯指背浮沉不定，而後可成。",
            ["巳_戌"] = "巳人生於戌日，日日東西，時難歸家，月德高照，雖多病但災難有極。",
            ["巳_未"] = "巳人生於未日，一生家庭欠安，凡事不順。",
            ["巳_申"] = "巳人生於申日，一生是非口舌，貴人得力可解。",
            ["巳_辰"] = "巳人生於辰日，一生喜事特別多，暈沉之病難免，美滿家庭，不安於內。",
            ["戌_丑"] = "狗人生於丑日，求謀不利，家庭糾紛，絲絲不斷，幸有女貴人化解之。",
            ["戌_亥"] = "狗人生於亥日，生意無聊，本去利空，太陽高照，進退自由，日日喜笑顏開。",
            ["戌_午"] = "狗人生於午日，家有財庫，外有白虎，守之不慎，必有大敗，但事業權位，必得人上。",
            ["戌_卯"] = "狗人生於卯日，異性多起爭奪，口舌破財，每遭不幸，月德照顧。",
            ["戌_子"] = "狗人生於子日，浮沉不定，呈若水上之舟，婚姻不遂，凡事欠順，每多困境。",
            ["戌_寅"] = "狗人生於寅日，時犯小人，好心為人，背後反遭人罵，凡事不遂心願，諸事不吉。",
            ["戌_巳"] = "狗人生於巳日，雖有官座暴敗，紫徽龍德二星高照，不致一敗如灰灰，喜色如流水。",
            ["戌_未"] = "狗人生於未日，口舌多，身體欠健康，禍頻生，尚有福星高照，遇難多救。",
            ["戌_申"] = "狗人生於申日，身騎天馬，奔馬東南西北，離鄉背井，終難歸家。",
            ["戌_辰"] = "狗人生於辰日，大耗破財，災難重重，多險惡風波。",
            ["戌_酉"] = "狗人生於酉日，病患沉，愁心不展，傷害難盡。",
            ["未_丑"] = "羊人生於丑日，家空手空，是非口舌之常有，宜謹慎小心。",
            ["未_亥"] = "羊人生於亥日，官符疊起，小人持多，住列三台，禦如流水。",
            ["未_午"] = "羊人生於午日，一生喜事多有，凡事順利，平坦樂道，前途光明。",
            ["未_卯"] = "羊人生於卯日，白虎臨命，財庫動搖，生財求謀注意小心，浮現之病難免，取官高尚，必掛將星。",
            ["未_子"] = "羊人生於子日，小破財，兼時犯小病，月德在命，臨危不救，酒色之春。",
            ["未_寅"] = "羊人生於寅日，命在紫徽星，諸事迪吉，喜氣重重，雖時有官非口舌，皆能逢凶化吉也。",
            ["未_巳"] = "羊人生於巳日，騎馬坐命，一發如山，求名求利，必是異鄉貴客，出外之人。",
            ["未_戌"] = "羊人生於戌日，太陽星照命，陰貴人必多，遇事不遂。",
            ["未_未"] = "羊人生於未日，顯貴聰敏，文章蓋世，鼇頭獨佔，時有利激之心，用功過甚。",
            ["未_申"] = "羊人生於申日，一生盡是紅鸞喜，太陽星高照，四路皆通，謀事高就，生意變通，能成天地空之感。",
            ["未_辰"] = "羊人生於辰日，福星高照，出外營謀，統率人群，權柄特高，是非口舌當見。",
            ["未_酉"] = "羊人生於酉日，災害難免。",
            ["申_丑"] = "猴人生於丑日，月德照命，凡事出一帆風順，出馬離鄉，浮沉不定。",
            ["申_亥"] = "猴人生於亥日，太陽星照命，多勾絞不明之麻煩，而無大困矣。",
            ["申_午"] = "猴人生於午日，凡事不吉，每多逆境。",
            ["申_卯"] = "猴人生於卯日，",
            ["申_子"] = "猴人生於子日，命帶將星，職躍虎門，小人四起，官府糾紛。",
            ["申_巳"] = "猴人生於巳日，天德福星照命，凡有是非口舌，當能逢凶化吉，事無困境。",
            ["申_戌"] = "猴人生於戌日，窀穸浮沉，不能安靜。",
            ["申_未"] = "猴人生於未日，單身從影，災難常起，身體欠康。",
            ["申_申"] = "猴人生於申日，紅鸞喜滿門，作福作壽，添子應孫，交友不利，頭疼不少。",
            ["申_辰"] = "猴人生於辰日，白虎坐命，破敗重重，但有天之禍，文才兼優，靈機應變力至足。",
            ["申_酉"] = "猴人生於酉日，桃花坐命，迎新送舊，情感紛紛，空空如也。",
            ["辰_丑"] = "辰人生於丑日，口舌多見，多有福星高照，天賜福德。",
            ["辰_亥"] = "辰人生於亥日，夫妻賢美得助，諸事順遂，逢凶化吉。",
            ["辰_午"] = "辰人生於午日，諸事逆境，浮沉不定，有如海中行舟。",
            ["辰_卯"] = "辰人生於卯日，混沌不堪，事無頭緒。",
            ["辰_子"] = "辰人生於子日，命帶將星，事業一帆風順，必有青雲之路，治國則國泰民樂，財豐利厚。",
            ["辰_寅"] = "辰人生於寅日，必出遠方為事，志在四方。",
            ["辰_巳"] = "辰人生於巳日，太陽星高照，謀事諸般吉，一生喜事重重，晦氣也多。",
            ["辰_戌"] = "辰人生於戌日，財庫欠穩。",
            ["辰_未"] = "辰人生於未日，太陽星高照，須防災難、煩惱之事。",
            ["辰_申"] = "辰人生於申日，凡謀必就，時有小人作弄。",
            ["辰_辰"] = "辰人生於辰日，華蓋坐命，學藝聰敏，時有鬱悶不開。",
            ["辰_酉"] = "辰人生於酉日，雖多小病，但月德高照，一切無妨。",
            ["酉_丑"] = "雞人生於丑日，一生多犯人，事職浮沉，一起一落，當見少吉。",
            ["酉_亥"] = "雞人生於亥日，重財輕義，發如猛虎，敗如浪沙。",
            ["酉_午"] = "雞人生於午日，喜氣極多，福星高照，凡事少憂。",
            ["酉_卯"] = "雞人生於卯日，大破大耗，災害不已，家逢風浪，能致掃空。",
            ["酉_子"] = "雞人生於子日，一生喜事重重，時防暴凡，勾絞麻煩之事，凡經營求謀欠利。",
            ["酉_寅"] = "雞人生於寅日，病患纏綿，幸有月德照臨，臨危有救也。",
            ["酉_巳"] = "雞人生於巳日，白虎掃財庫，家基當小心，災禍起伏，波折常逢。",
            ["酉_戌"] = "雞人生於戌日，人喜奉承，愛虛榮，好奢華，慷慨好施，惟晦氣不了。",
            ["酉_未"] = "雞人生於未日，相聚敦厚，性情和平誠實，間身只影，清寒淡薄。",
            ["酉_申"] = "雞人生於申日，膽大性暴，肆無忌憚，浮沉暈弱，宜忍慎。",
            ["酉_辰"] = "雞人生於辰日，紫徽星高照，諸事必順，求謀多遂，即時成功，天難時困，德恩保之。",
            ["酉_酉"] = "雞人生於酉日，體態豐厚，溫和忠恕，財路達，利恒旺，終身名利富有。",
        };

        // Ch.7 年生肖×時支（逐時考）
        var nayinCh7 = new Dictionary<string, string> {
            ["丑_丑"] = "牛人生於丑時，聰敏多少，文章蓋世，逢傷必淚。",
            ["丑_亥"] = "牛人生於亥時，飛天入地之能，多外少家，遠鄉作品。",
            ["丑_午"] = "牛人生於午時，桃花坐命，常欲迎情。",
            ["丑_卯"] = "牛人生於卯時，災難重重，凡事小心。",
            ["丑_子"] = "牛人生於子時，花常早拜，事業無愁，凡謀必成，子嗣旺相，一生逍遙，晚景甚佳。",
            ["丑_寅"] = "牛人生於寅時，花燭迎人，很少自由，怨天尤人，太陽高照，遇難有救，時欲單身為事。",
            ["丑_巳"] = "牛人生於巳時，官符常見，交友指背，財庫甚足，名旺他人。",
            ["丑_戌"] = "牛人生於戌時，每多口舌，流落異梓，驛馬為人，福星高照，到處可成。",
            ["丑_未"] = "牛人生於未時，丑未一沖，破碎牢災，厄運多走，一生少幸福。",
            ["丑_申"] = "牛人生於申時，天地賜福，順水行舟，如月之恒，如日這升，貴人得力，不謀自取，紫徽高照，不怕天災水患。",
            ["丑_辰"] = "牛人生於辰時，諸事人吉，流連煩繁，陰氣重疊。",
            ["丑_酉"] = "牛人生於酉時，白虎破財，浮沉多病，應防跌災。",
            ["亥_丑"] = "豬人生於丑時，言行不正，重同輕義，陰霧一生。",
            ["亥_亥"] = "豬人生於亥時，命犯太歲，事業浮沉不定，交友多犯指背星。",
            ["亥_午"] = "豬人生於午時，官星得位，漸入佳景，名旺財衰。",
            ["亥_卯"] = "豬人生於卯時，可為大事，人喜奉承，愛慕虛榮。",
            ["亥_子"] = "豬人生於子時，為人面貌清秀聰明才智敏多能，招蜂引蝶。",
            ["亥_寅"] = "豬人生於寅時，性情兇暴，藐視他人，大事難為。",
            ["亥_巳"] = "豬人生於巳時，驛馬坐命，發如猛虎，高低不准，晚景平常。",
            ["亥_戌"] = "豬人生於戌時，孤曩常遊，春風滿面，喜事頻繁。",
            ["亥_未"] = "豬人生於未時，白虎破財，雖聰敏能幹，天難難避。",
            ["亥_申"] = "豬人生於申時，花巧多計，多藝多能，辨論他人，名揚四海。",
            ["亥_辰"] = "豬人生於辰時，",
            ["亥_酉"] = "豬人生於酉時，本屬大旺，因酉亥相形，反遭破碎意外之事。",
            ["午_丑"] = "馬人生於丑時，",
            ["午_亥"] = "馬人生於亥時，亥水克置午火，病輕而難療，但月德臨命，重亦何防。",
            ["午_午"] = "馬人生於午時，午火比肩，家庭和睦，骨肉相親，命運豐隆，既富且壽，既安且寧，事職專權。",
            ["午_卯"] = "馬人生於卯時，卯木生持午火，利路宏路，滿堂春色，交朋接友多貴。",
            ["午_子"] = "馬人生於子時，子水克制午火，命逢紫徽星，諸謀大成，逍遙快樂，雖有悶氣小敗，不在其話下。",
            ["午_寅"] = "馬人生於寅時，寅木生持午火，雖時難事逆，交友不力，終有成功之日。",
            ["午_巳"] = "馬人生於巳時，巳火午火比肩。辛勤勞碌，越陌浮沉，事業不定。",
            ["午_戌"] = "馬人生於戌時，戌土行持午火，文章落選，官鬼小人，作亂複雜。",
            ["午_未"] = "馬人生於未時，未土行持午火，衣冠不下，財謀稱意，當有苦悶怨聲之感。",
            ["午_申"] = "馬人生於申時，申金克置午火，理直氣壯，職事必高，財散人離。",
            ["午_辰"] = "馬人生於辰時，午火生持辰土，暗淡清閒，浮沉不定，寂寞雙影，事不遂心，每有挫折。",
            ["午_酉"] = "馬人生於酉時，酉金克置午火，家運昌隆，喜多樂餘，不是作福做壽，就必應子應孫。",
            ["卯_丑"] = "卯人生於丑時，必是遠鄉客，時有不明不解之事。",
            ["卯_午"] = "卯人生於午時，夫妻賢美，有難得救，臨危無事耳。",
            ["卯_卯"] = "卯人生於卯時，大材大用，千里威聲，名揚四海。",
            ["卯_子"] = "卯人生於子時，迎新送舊，一片生機。",
            ["卯_寅"] = "卯人生於寅時，性燥心亂，必有亡神之處境。",
            ["卯_巳"] = "卯人生於巳時，命帶驛馬，終朝奔走四方，求名求利，孤單難歸。",
            ["卯_戌"] = "卯人生於戌時，紫徽星臨命，雖有王災不幸，終是逢凶化吉。卯人生於亥是，白虎星坐命，必有困苦這境，宜慎重為事耳。",
            ["卯_未"] = "卯人生於未時，聰敏出眾，文才兼能，飛來不測之事，宜慎重可也。",
            ["卯_申"] = "卯人生於申時，時渾沉不樂之病，但有月德照臨，幸無大害矣。",
            ["卯_辰"] = "卯人生於辰時，陰霧不晴，一時不得出頭天。",
            ["卯_酉"] = "卯人生於酉時，時有破碎之難，空虛之苦。",
            ["子_丑"] = "鼠人生於丑時，多招美貌佳人，常代女子掃財，宜女不宜男。",
            ["子_亥"] = "鼠人生於亥時，謹防身體。牛",
            ["子_午"] = "鼠人生於午時，多遭口舌。",
            ["子_卯"] = "鼠人生於卯時，習天膽大，有張飛之勇，雖不能登臺拜將亦成人間真人。",
            ["子_子"] = "鼠人生於子時，宜早完婚，免致後嗣空虛，事業一帆風順。",
            ["子_寅"] = "鼠人生於寅時，多招風流才子，常見男子掃財，宜男不宜女。",
            ["子_巳"] = "鼠人生於巳時，當代破碎。",
            ["子_戌"] = "鼠人生於戌時，不見其利，多在其危，宜善守之。",
            ["子_未"] = "鼠人生於未時，必成必敗，一反一覆，形似水浪。",
            ["子_申"] = "鼠人生於申時，一生潔吉，雖有小破財，不在其害。",
            ["子_辰"] = "鼠人生於辰時，難始終和一。",
            ["子_酉"] = "鼠人生於酉時，子女滿堂喜慶。",
            ["寅_丑"] = "寅人生於丑時，紅鸞喜星，為人善言聰敏，喜悅一方，老少合歡，雖有些小病，笑笑了之。",
            ["寅_亥"] = "寅人生於亥時，衰弱枝葉簿，雖有才能，一時榮華，終日有敗也。",
            ["寅_午"] = "寅人生於午時，膽才過人，為國棟臣，建大業奏奇功，時退如原。",
            ["寅_卯"] = "寅人生於卯時，花前月下常留影，談笑春風意多情，為人樂遊，一生少苦惱。",
            ["寅_子"] = "寅人生於子時，凡事注重。",
            ["寅_寅"] = "寅人生於寅時，聰敏至貴，學富五車，英傑才人。",
            ["寅_巳"] = "寅人生於巳時，陰霧不開，悶氣沉，纏身不脫。",
            ["寅_戌"] = "寅人生於戌時，華蓋坐命，臨機應變，奇巧才能，他日光門耀祖。",
            ["寅_未"] = "寅人生於未時，天送之喜，掛馬之名，雖有風波險境，月德照臨，逢凶化吉，百事可成也。",
            ["寅_申"] = "寅人生於申時，命坐驛馬，忠心報國。",
            ["寅_辰"] = "寅人生於辰時，才儲八半，少得貴人。",
            ["寅_酉"] = "寅人生於酉時，",
            ["巳_丑"] = "巳人生於丑時，心謀心成。雖有破敗，飛來意外之事，終有成功之日。",
            ["巳_亥"] = "巳人生於亥時，驛馬坐命，日日遠走異鄉，奪波四野。",
            ["巳_午"] = "巳人生於午時，交友益廣，逢人多情，太陽星高照，雖感晦氣，過日即除。",
            ["巳_卯"] = "巳人生於卯時，東跑西跑，無日不歸家。",
            ["巳_子"] = "巳人生於子時，必見天厄，在龍德保駕，雖危無大不害之事。",
            ["巳_寅"] = "巳人生於寅時，平意平常，凡事在人，成事在天，福星高照，前途光明。",
            ["巳_巳"] = "巳人生於巳時，多犯指背星，浮沉不定。",
            ["巳_戌"] = "巳人生於戌時，月德照臨，躍馬千里，時見小病及破財。",
            ["巳_未"] = "巳人生於未時，華蓋坐命，敏捷賢能，太歲犯上，浮沉之命。",
            ["巳_申"] = "巳人生於申時，貴人坐命，處處可行，但難免不平之事，麻煩不斷",
            ["巳_辰"] = "巳人生於辰時，一帆風順，喜事重重，不戀自愛，頭雅清尚，交友得益。",
            ["巳_酉"] = "巳人生於酉時，小人時有，中事登將星，財積如山。",
            ["戌_丑"] = "狗人生於丑時，小破小災當有，煩惱意外之事特多，尚有貴人保解其半，稍有平安。",
            ["戌_亥"] = "狗人生於亥時，單身流浪異鄉，但太陽星高照，前途光明。",
            ["戌_午"] = "狗人生於午時，職權在握，名揚四海，白虎臨命，途匱必破。",
            ["戌_卯"] = "狗人生於卯時，當有病患小耗財，卯與戍合，尚有月德在命，決無大難。",
            ["戌_子"] = "狗人生於子時，不順其時，諸謀不利，漂流不定，狼籍不堪。",
            ["戌_寅"] = "狗人生於寅時，身前身後，時應小心，晦氣不開，每遇挫折。",
            ["戌_巳"] = "狗人生於巳時，雖有無難，龍德之恩保釋，不致有暴敗。",
            ["戌_未"] = "狗人生於未時，孤單隻身，福星臨命，遇事無難，諸謀大吉。",
            ["戌_申"] = "狗人生於申時，躍馬千里之客，離鄉有期，事從心願。",
            ["戌_辰"] = "狗人生於辰時，財庫須動，月空破碎。",
            ["戌_酉"] = "狗人生於酉時，身體欠健，病患多累。",
            ["未_丑"] = "羊人生於丑時，普通行運，偶有破財，但人才高尚，凡事處理，有序無亂。",
            ["未_亥"] = "羊人生於亥時，小人頻繁，官府是非不斷，然有權位，即能免除。",
            ["未_午"] = "羊人生於午時，一身喜氣，玉堂常當，花花家園，樂意無邊，雖有小病晝夜即除也。",
            ["未_卯"] = "羊人生於卯時，命帶將星，熱力高強，浮沉不穩固。",
            ["未_子"] = "羊人生於子時，每有小耗破財，月德照臨，凡事遇難即解也。",
            ["未_寅"] = "羊人生於寅時，紫徽星高照，雖有官非晦氣，終是有救也。",
            ["未_巳"] = "羊人生於巳時，",
            ["未_戌"] = "羊人生於戌時，婚姻困擾不利，勾設貫孝事較多。",
            ["未_未"] = "羊人生於未時，畢蓋坐命敏敏賢能，太歲犯上，時見不刺激浮沉之癸。",
            ["未_申"] = "羊人生於申時，太陽呈高照，喜事重重，凡謀事利，雖孤單雙影，但少困難之境。",
            ["未_辰"] = "羊人生於辰時，福星照命，離鄉避祖，東去西行，南謀西就，求名求得，外出一生，盼能早歸。",
            ["未_酉"] = "羊人生於酉時，是非口舌，官府牢災，聊於常事。",
            ["申_丑"] = "猴人生於丑時，丑土生持申金，威聲稟稟，月德照命，四路可行，一帆風順。",
            ["申_亥"] = "猴人生於亥時，申金生持亥水，太陽照命，雖見官府暴敗，多得人和調理，凡事得順。",
            ["申_午"] = "猴人生於午時，午火克置申金，一身少吉，此命欠祥。",
            ["申_卯"] = "猴人生於卯時，申金克置卯木，天雖暴敗，但晨卯紫徽星照臨，逢凶化吉，轉為大喜。",
            ["申_子"] = "猴人生於子時，申金生持子水，事業必旺，但受小人之亂，職權克置難起。",
            ["申_寅"] = "猴人生於寅時，申金克置寅木，雖能出官為事，恰似水浮萍，有名無實。",
            ["申_巳"] = "猴人生於巳時，巳火克直申金，大小是非煩，但事業職位，遇得貴人協助，福星天德高照，有利無害。",
            ["申_戌"] = "猴人生於戌時，",
            ["申_未"] = "猴人生於未時，未土生持申金，陰蓋與陽，疾病難免，不得人和。",
            ["申_申"] = "猴人生於申時，喜氣洋洋，背人過消費品，交友傷神。",
            ["申_辰"] = "猴人生於辰時，辰為生持申金，才能過人，誚是諸凡順示，但是辰屬天羅，臨之破敗。",
            ["申_酉"] = "猴人生於酉時，喜遊街巷，晦氣不堪。",
            ["辰_丑"] = "辰人生於丑時，吉星高照，求謀大事，喜笑美滿，順時退之，而忍性做事，終日善矣。",
            ["辰_亥"] = "辰人生於亥時，大吉大利，德星高照，逢凶化吉，一生少苦惱，凡事多利矣。",
            ["辰_午"] = "辰人生於午時，必暴狂躁，浮沉漂泊，時見跌撲，凡事宜慎重，則家運可順也。",
            ["辰_卯"] = "辰人生於卯時，生性渾沉，需用力為事，才能成功。",
            ["辰_子"] = "辰人生於子時，事業順利，大展鴻圖，諸事吉順，終可成功。",
            ["辰_寅"] = "辰人生於寅時，天南地北，終日少時歸家，命犯天狗星，凡事小心。",
            ["辰_巳"] = "辰人生於巳時，百事可取，雖有災禍，無大凶難，男女不是填房過繼，便有兩家之春。",
            ["辰_戌"] = "辰人生於戌時，時有狼狽之事，恐防官非口舌，時有小破財，凡事小心為利。",
            ["辰_未"] = "辰人生於未時，流沛麻煩之事特多，雖有破敗，亦有成功的機會。",
            ["辰_申"] = "辰人生於申時，財豐利足，四方皆利，官鬼頻繁，官符常見，宜謹慎從事。",
            ["辰_辰"] = "辰人生於辰時，聰明主貴，衣食豐足，凡謀得志，鄉里和睦，親友稱道。",
            ["辰_酉"] = "辰人生於酉時，桃花坐命，慎防傾倒，防是非口舌，雖有財利也有驚憂。",
            ["酉_丑"] = "雞人生於丑時，有受眾人恭敬多有小人暗算。",
            ["酉_亥"] = "雞人生於亥時，離鄉背井，財利路達，但終身孤影。",
            ["酉_午"] = "雞人生於午時，聰敏多能，風流灑色，口舌多起，但事業有成。",
            ["酉_卯"] = "雞人生於卯時，性情聰敏而怪異，反面大耗，是非口舌，一生頻繁。",
            ["酉_子"] = "雞人生於子時，繁榮滿堂，主人懦弱不任大事，反為小人致敗。",
            ["酉_寅"] = "雞人生於寅時，忠厚誠實，性情無准，好壞反復。",
            ["酉_巳"] = "雞人生於巳時，精神發越，然施威逞才，一生之中實難頻繁。",
            ["酉_戌"] = "雞人生於戌時，終身出外為事，愛情多晦氣。",
            ["酉_未"] = "雞人生於未時，兇暴勢強，奪財滋禍。",
            ["酉_申"] = "雞人生於申時，身強材弱，凡事少利。",
            ["酉_辰"] = "雞人生於辰時，性情敦厚，言行相顧，職權至高，終途暴敗。",
            ["酉_酉"] = "雞人生於酉時，官祿至貴，財豐利達，浮沉小疾難免。",
        };

        var sb2 = new System.Text.StringBuilder();

        // === 年柱納音 ===
        string yNayin = nayin60.GetValueOrDefault(yStem + yBranch, "");
        if (!string.IsNullOrEmpty(yNayin))
        {
            sb2.AppendLine($"【年柱納音：{yNayin}】");
            // Ch.1: 五行命名原理
            if (nayinCh1.TryGetValue(yNayin, out var yDesc1) && !string.IsNullOrEmpty(yDesc1))
                sb2.AppendLine(yDesc1);
            // Ch.4: 生肖+納音特質（不重疊）
            if (nayinCh4.TryGetValue(yStem + yBranch, out var yDesc4) && !string.IsNullOrEmpty(yDesc4))
                sb2.AppendLine(yDesc4);
            // Ch.3: 年命推算走向（取不重複段落）
            if (nayinCh3.TryGetValue(yNayin, out var yDesc3) && !string.IsNullOrEmpty(yDesc3))
                sb2.AppendLine(yDesc3);
            sb2.AppendLine();
        }

        // === 月柱：年生肖×月支 ===
        string ch5Key = yBranch + "_" + mBranch;
        if (nayinCh5.TryGetValue(ch5Key, out var mDesc) && !string.IsNullOrEmpty(mDesc))
        {
            sb2.AppendLine("【月柱生肖論】");
            sb2.AppendLine(mDesc);
            sb2.AppendLine();
        }

        // === 日柱納音 ===
        string dNayin = nayin60.GetValueOrDefault(dStem + dBranch, "");
        if (!string.IsNullOrEmpty(dNayin))
        {
            sb2.AppendLine($"【日柱納音：{dNayin}】");
            // Ch.2: 日柱納音細注
            if (nayinCh2.TryGetValue(dStem + dBranch, out var dDesc2) && !string.IsNullOrEmpty(dDesc2))
                sb2.AppendLine(dDesc2);
            // Ch.6: 年生肖×日支
            string ch6Key = yBranch + "_" + dBranch;
            if (nayinCh6.TryGetValue(ch6Key, out var dDesc6) && !string.IsNullOrEmpty(dDesc6))
                sb2.AppendLine(dDesc6);
            sb2.AppendLine();
        }

        // === 時柱：年生肖×時支 ===
        string ch7Key = yBranch + "_" + hBranch;
        if (nayinCh7.TryGetValue(ch7Key, out var hDesc) && !string.IsNullOrEmpty(hDesc))
        {
            sb2.AppendLine("【時柱生肖論】");
            sb2.AppendLine(hDesc);
        }

        return sb2.ToString().Trim();
    }


        private static string LfShiGanXiangFa(string dStem, string mBranch)
        {
            var stemBase = new Dictionary<string, string>
            {
                ["甲"] = "甲木之人有名望，正直，說話有力度，在山上為樹，在市井為木。",
                ["乙"] = "乙木象花草藤蔓，隨風飄動但根不動，曲直蜿蜒，柔韌善變通，能在各種環境中找到生存之道。",
                ["丙"] = "丙為陽火，代表熱烈光明，愛出風頭，不計後果，心直口快，是個熱心腸，愛慕虛榮，事業方面與文化有關係。",
                ["丁"] = "丁火代表理想願望，形態尖銳，為人比較專一，辦事有頭腦，認真思考，性格表面柔和和順，非常有心機，處事比較細膩，旺相時比較剛烈，事業主要代表文化技巧。",
                ["戊"] = "戊土就像城牆的外表、江河的堤壩，外表堅實，擋風擋水，但是一定要根基重，否則容易城倒堤崩。象法就是外表的保護包裝，同時也是一種障礙，土多必須要水來通，可用甲木來制約。",
                ["己"] = "己土是陰土，是平整的大地，小的來看是莊園、庭院，己土胸懷很大，能容納萬物，個性內向含蓄謹慎，容納力強同時也多疑，喜歡管理，不喜歡強制性約束，喜歡先謀劃後行事，愛使計謀，說話不直接傷人，拐彎抹角使對方滿意。",
                ["庚"] = "庚金屬於礦石，含金量很高的礦石，表示本人的技術性，性情方面都很專一。為人本性非常仗義，處事態度屬於強硬型、命令型的，干什麼事業只要認准了，誰的話也聽不進去，金頑的時候喜丁火。秋天金最旺，喜見火，可以受約束，但絕對不能見水，水見金沉。",
                ["辛"] = "辛金屬陰金，成品的金屬，帶尖帶刃的，也代表金銀首飾，比較名貴。辛金愛出頭，愛美，好炫耀，辦事比較善變，說話比較尖銳帶刺，做事精打細算，思想比較超前，有個性，點子比較多，經常做些出乎意料的事。",
                ["壬"] = "壬水陽水屬於大水，特性代表聰明智慧，變化多端，應變能力強，一般職業代表運輸、流通、物資集散。地支聚成水勢，非漂流在外、外地發展，本地是起不來的。",
                ["癸"] = "癸水為陰水，濕氣，看不到的水，代表困難閉塞，就需要流通，流通了就有了智慧，就能脫困，但是也不能過旺，則會出現錯誤。思想比較廣泛，非常小心細膩，多疑，癸水多算計，抗壓能力強。"
            };

            var monthDesc = new Dictionary<string, string>
            {
                // === 甲木 ===
                ["甲寅"] = "春天甲木，狀態為生機盈然，含苞待放，喜陽光之照耀，不喜見水，不喜見庚金，見辛金為修剪之象，把枝枝叉叉修剪之後更利於生發。行事半遮半掩，小心翼翼，不是一步到位。",
                ["甲卯"] = "春天甲木，不喜見土，見土克財克父，婚姻易有波折，見辛金為修剪之象，把渾身的毛病去掉了更利於生發，因為二月甲木比較任性，毛病多，而見水還是有些爭寒。",
                ["甲辰"] = "春天甲木末期，見辛金為修剪之象更利生發，不喜水爭寒，不喜庚金強剋。天干透出丁火己土對其有利，不要透出乙木，透出乙木必克妻，無此因素之人經商有能力，財運好，主意正，有能力有魄力。",
                ["甲巳"] = "四月甲木，此時陽光充足，樹葉繁茂，外表道貌岸然，即把自身的缺點掩蓋起來。怕火，喜水，不喜在陽光下而喜在陰暗中生存。女命生在四月較好，如果不見官的命中說明其聰明，口齒伶俐；男命則內心能量需注意正向發揮，愛表現自己，愛出風頭。",
                ["甲午"] = "夏天甲木，陽光充足，樹葉繁茂，外表道貌岸然，善掩蓋自身缺點，怕火，喜水，更宜在陰暗中生存，需注意內心能量的正向發揮。",
                ["甲未"] = "六月份甲木，六月為未土，墓庫身庫，甲木通強根。天干透出丁火，己土對其有利，不要透出乙木，透出乙木必克妻，無此因素之人經商有能力，財運好，主意正，有能力有魄力，這個甲木能體現自身的能量出來。",
                ["甲申"] = "秋天甲木最次。秋風掃落葉，偽裝無，內心空，剩外表虛假軀干。不宜見金，喜見水，有水滋養，甲木能量則有，如無水再見土，此類人則是一具軀殼。",
                ["甲酉"] = "秋天甲木最次。秋風掃落葉，偽裝無，內心空，剩外表虛假軀干。不宜見金，喜見水，有水滋養，甲木能量則有，如無水再見土，此類人則是一具軀殼。",
                ["甲戌"] = "秋天甲木最次。秋風掃落葉，偽裝無，內心空，剩外表虛假軀干。不宜見金，喜見水，有水滋養，甲木能量則有，如無水再見土，此類人則是一具軀殼。",
                ["甲亥"] = "冬天的甲木，已經成材了，喜見到庚金就成材了，成為棟樑之才，有用之材，見庚辛金在仕途，經商方面都有發展，有能力，是成功人士，不論男女見庚辛金都會有發展。",
                ["甲子"] = "冬天的甲木，已經成材了，喜見到庚金就成材了，成為棟樑之才，有用之材，見庚辛金在仕途，經商方面都有發展，有能力，是成功人士，不論男女見庚辛金都會有發展。",
                ["甲丑"] = "冬天的甲木，已經成材了，喜見到庚金就成材了，成為棟樑之才，有用之材，見庚辛金在仕途，經商方面都有發展，有能力，是成功人士，不論男女見庚辛金都會有發展。",
                // === 乙木 ===
                ["乙寅"] = "正月乙木，天干別透甲，喜見傷官財官。男命逆行運不好，女的順行運好，多福祿，天干勿透甲，否則克妻財。喜見傷官或七殺，庚金制頑，不喜見水，喜丙火驅寒，而丁火為煙火暖不了土作用不大，喜戊土，因為戊土能合住癸水去寒為吉利。",
                ["乙卯"] = "二月乙木，喜見財官，地支見火金，白手起家，但無依靠，天干地支見金且有根，為自身勞動所得。見庚辛金，木最強，但不是棟樑之材，做不了一把手，花草只是裝飾，副手秘書，輔助領導。",
                ["乙辰"] = "辰月乙木，三月濕土，肥土，雜氣，乙木紮根在肥土之上，走西方運，八字見金，運勢不錯，怕戌。地支辰戌沖，多顛倒，再逢刑傷，壽不高，財不好。",
                ["乙巳"] = "巳月乙木，立夏，東北運好。前提地支對乙木有生扶，有根基，否則無依靠，身體瘦弱則不好。木盛見丁火辛金，七殺有制化為權，平步青雲發少年，在武職方面有造就。",
                ["乙午"] = "夏乙木，巳月立夏，風吹草動，草木出土，繁茂，風一吹且日月哺育，成長更加快速。枯草落地為母，化為肥滋養，寒意全無。不能見丁火，丁火是煙火，煙火把綠色草木的根本傷及，此格局之人短壽。",
                ["乙未"] = "乙木只有生在秋天才可從可化。夏乙木性喜繁茂，不宜見丁火煙火，宜火土調候，戊土最好。",
                ["乙申"] = "申月乙木，乙木見水而盛，見水好，走水土運，八字見水土有名望，不見得有權力。無水見金，從化還好，如果水弱見金，從也從不了，生存很困難。",
                ["乙酉"] = "酉月乙木，喜火來攻金，忌土泄火生金，乙木地支有根，干透庚辛金，特別是辛金，無火來攻，身體婚姻事業都不太順，特別是辛金要注意。",
                ["乙戌"] = "戌月乙木，喜金水，不喜見土火木，（不合化前提下）化了另論。乙木在秋天方可從可化，戌月是秋天尾聲。",
                ["乙亥"] = "亥月乙木，乙木需培根養蓄，火土驅寒，戊土最好。葉落歸根，根向下紮，只宜陽光，不喜陰，不能見水，此時見水，沖亂草的規則。",
                ["乙子"] = "子月乙木，冬月，原命中不能見七煞，辛金為七殺，因為辛金生寒，金水增寒。男命逆行運不佳，無官無殺最好，不能見正偏印（水）。",
                ["乙丑"] = "丑月乙木，運行東方運，喜見丙丁，不富則貴，一陽驅寒，丙克庚，丁克辛，七煞有制。",
                // === 丙火 ===
                ["丙寅"] = "正月丙火，長生之地，喜南方火運，木運也可以，但不能太旺，否則火被晦了，木多了喜見庚金，財官雙得；不喜見辛金，容易貪戀美色，玩物喪志，火多喜壬水，劫財多了喜癸水，金多喜見木，不喜見丁火爭輝。",
                ["丙卯"] = "二月丙火，弱的話喜見比扶，還不喜見丁火，見財雖得，但不大不多，且運走財隨得但不存，因為二月財為絕地，木多見煞不吉，喜經商投資，丙火有點急性，不考慮後果，不聽勸，總之木多了喜見金。丙火生在二月不論強弱，都有感情的風波，因為感情太豐富，見一個喜歡一個。",
                ["丙辰"] = "三月丙火，辰特殊的一個月，濕土雜氣，透土喜見財，富裕之人，不喜見官煞，因為辰本身為水庫，火就起不到任何作用了。困難時能夠齊心合力，一但有財就會起爭劫而傷情。",
                ["丙巳"] = "四月丙火，為見祿，喜見煞，丁火多了喜見官，丙火多了喜見煞，但是男命不喜見戌運，火一入庫必有災禍，其他運還可以，巳火入庫必有災禍。",
                ["丙午"] = "五月丙火，最旺，陽運當令，無論強弱，不喜見子，陽運逢沖必有災，陽運最怕沖，有根有氣，有災無禍，無氣災禍並發。",
                ["丙未"] = "六月丙火，傷官當令，喜見印星，正印和傷官並見必出文化人，很聰明，見水不吉，水重走東南方運，最忌諱官煞混雜，必有災禍，不混雜尚可。",
                ["丙申"] = "七月丙火，偏弱可從可化，沒有根氣，化財了，如果中和，不弱不強，見水必傷，走東南方運都會有成就。",
                ["丙酉"] = "八月丙火，火還有點餘氣，偏弱，喜見印比，運行東南方還可以，但終究沒有大的氣勢，如果火弱極了，從了酉金了必出奇人奇事，如果從了再走東南方運，那就破格了，破格背祿必有凶禍。",
                ["丙戌"] = "九月丙火，食傷當令，火有庫根強根，見官或者見煞都出貴，如官殺混雜必剝祿，能當官但終必被罷官，且見財因財而招災。",
                ["丙亥"] = "十月丙火，付出多不得利而傷了元氣，大運走東南運出貴，走西北運，壽不長，運不佳，身體不佳，有印星佳，財官雙得，沒有印星，從又從不了就出問題了。",
                ["丙子"] = "十一月丙火，喜見木，東南運發達，出貴不出富，走西北敗運，且眼睛有病，盲目前行，必出大亂。",
                ["丙丑"] = "十二月丙火，付出多，收穫的也多，比較累，土多配印財必得，因為丑為金庫，怕走亥運，水旺滅火必見災。",
                // === 丁火 ===
                ["丁寅"] = "正月的丁火，為正印，不見庚甲，劈不了木，就生不起來丁火，見金見甲最吉，如有壬水喜丁來合，這時就不喜再見金了，因為見金會生水就不會去劈甲了。木多水生，變成濕木晦火之象，反而不利丁火。這個月的丁火是帶有希望的，希望有乾柴生輔，那麼見財得財，見官得官，歲運見之得之。",
                ["丁卯"] = "二月的丁火，為花草初生的時候，原來的乾草可以生丁火，為偏印，不可以見水，特別是地支當中見水成河，必有災。這時也可以見點水，但不能成河成勢，可以獨官獨殺，利仕途發展，但不可官殺混雜。可以助身，為人事業多，思維比較細膩，心思比較重，看事情也比較尖銳。",
                ["丁辰"] = "三月的丁火，丙火出來的早，奪丁火的光，作用開始減弱，喜火勢幫扶，增強自己的生存能力。等丙火出來丁火就沒用了，可以轉化為爐火，可以煉鋼煉鐵了，不喜木勢，忌地支成河，干透壬水也可以，可從官從殺。",
                ["丁巳"] = "四月丙火最旺，最忌丙火出現，這個時候，如果沒有水就過旺過強了，如果見壬水，可以制約丙火，又能合丁火，再見西方金，怕見甲見丙不好。",
                ["丁午"] = "五月為見祿火，火過旺而不利，喜水來平衡控制，否則必有災禍，火可以燎原了，烤光了，把金都煉成水了，這時候就喜水土之運了，土可以泄火，水可以制約火。",
                ["丁未"] = "六月為爐中火，得令而有燥性，失去了細膩的性格。喜見金水，火旺可見金而成氣，見水控火可成勢，怕見木。",
                ["丁申"] = "七月的申金當令，正財，秋老虎，喜木火運，忌癸水，喜見壬水，喜官不喜殺，丁火已失去頑劣之氣，喜見甲乙木，庚能劈甲，乙可以得到修剪枝葉，火就得了勢，可以煉金，一但水多成勢，生命就有危險了。",
                ["丁酉"] = "八月偏財，丁火可借水木之勢制頑金。酉金是晚上5點到7點，八月這時候正好該開燈了，太陽也落山了，無力爭輝，丁火起用了，丁火長生在酉就是這個道理。",
                ["丁戌"] = "九月戌傷官月，人們都回家了，也點上燈了，家裡一片光明，生在戌月不錯，人很有愛心，喜走東南運，見木火有成就，見水可以但不能成勢，成勢就滅火了。",
                ["丁亥"] = "十月亥月為正官月，丁火喜得氣勢，有火勢，有木氣，十月的木為燥木，所以能生火，福氣不錯，如果亥子丑會成水局，從化不得，反成災禍，不從不化可定災禍連連，克制太重。",
                ["丁子"] = "十一月子月七殺當令，雖然殺旺，如見庚見甲，庚金劈甲，形成木火通明，這樣的人更顯貴，七殺有制化為權，平步青雲發少年。走木火或者火土運都可以。天干透庚金，但地支也不能形成申酉戌的局，否則壽不長，身體欠佳，沒有依靠，沒有後代，生女不生男。",
                ["丁丑"] = "十二月丑月，土金旺，土重滅火，金重而晦火，見甲乙木或丁火幫扶，有生存之機，但是生在丑月的丁火命運都不是太好，一般平平。這時候喜甲乙木，乙木成乾草可以生火，甲木可以劈甲生丁了。",
                // === 戊土 ===
                ["戊寅"] = "正月寅月當令，七殺當令，沒有印星化殺，氣不通，沒有火的力量終不能成器，原局沒有火，大運行火也是平平，木重再行東方運，壽短無後，七殺過重了，見火化殺，氣成通，無需水來通。原局有火化殺，不喜見水了，所以原局有火不能走北方運。",
                ["戊卯"] = "二月卯木正官當權，如原局正官配印行東西南北運皆吉，但局中辛酉金不可多見，多見母親必短壽。見水，身難逃災，自身難逃，火勢水局而大富，水勢火局而貴，火有勢水有局必是大富之人，水有勢火有局必出貴。",
                ["戊辰"] = "辰月戊土非常強盛，喜見財官，但是地支不能見三庫，辰戌丑未見了三個必有牢獄之災。水木之勢多生子，火土之局少生男，大運行東北有利，行西南蹉跎，運勢不佳。",
                ["戊巳"] = "四月見祿，偏財七殺為喜，更喜西北去尋木，又走西南多不利，這個月喜水潤，喜木制都是有功。",
                ["戊午"] = "五月的印當權，柱中喜見財，最忌火成勢，印多而梟食，必懶惰，每個命局中只要是偏正印多必是懶惰之人。見金水調劑燥土烈火，見木更忌，木雖能制土，但更能生火。官生印，印生身，這種氣過旺了。",
                ["戊未"] = "六月日干強盛，喜殺帶根，甲木根在未，燥土飛揚，喜走水木之鄉，見火土而遭殃，不可見三庫丑戌，見之有牢獄之災，強旺有子，有子也不得力，為不孝之子。傷官主勢泄身太過，有炎火才能有利，在文學印刷包裝土建發展。",
                ["戊申"] = "七月申金食神當權，不喜財殺，這時候以食傷為財最佳，食傷代表技術藝術，一行火土之鄉，多主智慧，能做一首好文章，是個文人，為人聰明，但是不伶俐。",
                ["戊酉"] = "八月是傷官主事，多主經商、運輸，主文章主文印，若無財殺，八字也平常，運行東北財官之鄉名聲必響，如果行西南方則不利。",
                ["戊戌"] = "九月比肩，為火庫，日主強，原局財殺得氣得力是最吉祥的了。",
                ["戊亥"] = "十月偏財主勢，再有金水幫扶，都是父輩強，地支火勢喜財殺，最怕地支一片汪洋，戊土雖喜水來通，但也不喜成汪洋，一片水局。地支火局能暖水生土，但不喜再走北方運，總之見火為吉，見水為凶。",
                ["戊子"] = "十一月正財當令，地支見水見火，最怕殺傷位，七殺傷官之位，有火可以暖局，土可以助身制水，怕木來克，金生水旺，穿透了戊土，牆倒堤崩。戊生子月，寒氣盛見丙火，怕見金水木破局。",
                ["戊丑"] = "十二月丑月，寒氣太盛，有丙火，喜見財殺，沒有丙火，忌見財殺。有了丙火，忌金，更忌三庫臨土，牆倒堤崩，這樣都有意外之災，在哪個柱上哪個柱應驗。",
                // === 己土 ===
                ["己寅"] = "正月己土，春寒水冷，己土能有什麼作用呢？什麼也生長不出來，喜暖，喜溫室，就可以發芽生長，怕走東方七殺運，那命運是多苦多難，能夠含悲強忍，將來走出火運一定會發達。如果強勢出頭，必敗無疑，甚至危及性命，在沒有輔助情況下，是一個積弱的狀態，經受不住壓力，必倒無疑。",
                ["己卯"] = "二月的卯木，為偏官，七殺當令，無生無依，可以棄命從殺，如果有生又有依靠，氣勢得利，勢力均衡，再走南方火運就會有發展，財官雙得。這樣的命運要比無力強的多。",
                ["己辰"] = "三月辰月當權，氣勢比較真，濕土得到輔助，又當令，運勢不管順逆財官易得，家和齊美。得子也是顯貴的，己土生四季相對比較好，辰月中有財有官還能幫扶。",
                ["己巳"] = "四月己土，身強，必成燥土無疑，特別是丁火偏印。女命見了甲木或者乙木都不算好，傷官見官為禍百端，女的走正偏官運都不好，男的也會有波折。",
                ["己午"] = "五月己土最忌丁火，形成燥土，必須原局有水還要通根，無水有金，多屬文書少得財，喜水潤土而滅火，喜官不喜殺，見祿格走旺運，因為過旺必有衰，忌官殺混雜，男的牢獄之災，女的多婚之象。",
                ["己未"] = "六月為四季土，裡頭沒有水是燥土。未土當中暗藏玄機，喜水來潤，和五月區別不大，見甲逢合比較好，見午逢合也不錯，合住燥土，不至於燥土飛揚，這種命局行至中年必發達。",
                ["己申"] = "七月申金傷官主事，身輕事不安，日主一弱命運不安寧，最怕寅卯辰殺局，必遭官非，瑣碎事也多，女多婚，男牢獄，申子辰合水局泄掉了傷官，天干再透出正印，傷官配印，文章必顯，出文化人。",
                ["己酉"] = "八月酉金食神當令，身弱命不牢，沒有生扶，身體弱而多病，那麼要衰就從化為好，有生扶行南方運個人雖好，但整個大局來看也是平平而已。",
                ["己戌"] = "九月戌月火出來了，己土生四季，四季有分別，土旺火旺，無論順逆都很平穩，不會有大災大難，行財官運必發達，怕行東南運也不是太好，東南方內有辰土，辰戌相沖，坐根被沖，禍事連連。",
                ["己亥"] = "十月亥水財當令，但不一定就能得到財，財旺代表父親，月令也代表父宮，也代表兄弟宮。那麼地支當中形成寅午戌火局，巳午未會局，不富則貴；反過來形成了水局水勢必遭災，身弱見財而生災。",
                ["己子"] = "十一月子月，屬於偏財，見獨官獨殺為清，如果官殺混雜無論男女必遭災，最怕日主太弱，再走寅卯月，必凋零，十一月為凍土，在上面栽樹種地能成麼？行南方火運佳。",
                ["己丑"] = "臘月己土，也喜南方火運，有火能驅寒，沒火即使有官有殺也當不了權，說了也不算，還是喜印星生暖方為吉，南方運來輔為佳，官制劫財妻得救，殺制比肩兄安寧，沒有南方火，官殺什麼都沒用，無所事事，默默無聞，有了南方火，文章顯赫，名聲大振。",
                // === 庚金 ===
                ["庚寅"] = "正月的庚金，日干比較弱，在絕地，天干透土生金，有火煉金，庚金不怕火，因為正月的木旺，生火，泄木之頑氣，可以暖金，火生土，土生金，形成一個迴圈。",
                ["庚卯"] = "二月庚金，喜官不喜殺，因為丁火屬於明火，明火煉金可成氣，丙火燥金奪光輝，沒有土再見官殺，這個人的運氣不太好，非常不順。",
                ["庚辰"] = "三月庚金，有土來生，但土不能太多，土多金埋，沒什麼事可干，再有多大的能力也使不出來，土多用木。見水金沉，土也混，就更糟糕了。",
                ["庚巳"] = "四月庚金，七殺強，奪金光輝，喜殺印相配，七殺有制，如果金沒根，殺也沒制，則一生多磨難。",
                ["庚午"] = "五月庚金，火為明火，非常高興，再見水，更有權利了，但水不能太大，有主意，說話比較強盛，走東南方運，見財得財，見官得官，有財官之喜。",
                ["庚未"] = "六月庚金，燥土不生金，六月土為飛揚之土，為灰塵，喜見水，不能見火了，否則土更燥，喜水木運，庚金有權利。",
                ["庚申"] = "七月庚金，太鋼太強，更喜火來制約，見木得富，見火得貴，多生子少生女，見財官大運都佳，忌土金，這時候有能力發揮不出來。",
                ["庚酉"] = "八月庚金，更強了，這時候屬肅殺，喜火，有火必有權利，有生殺之權，有財無官殺，這叫從格了。",
                ["庚戌"] = "九月庚金，喜東方財運，丙丁火透出天干出貴，但不能官殺混雜，原局有丙丁火且通根，再走出丙丁火運，這叫官殺混雜身有災。",
                ["庚亥"] = "十月庚金，有點蕭條了，見土為妙，喜走土金運，發富貴，但多生女少生男，走水運、火運，原局有毛病肯定會有災的。",
                ["庚子"] = "十一月庚金，金見水沉，喜走土制水，而生金，走出印行，或走正官運，八字還不錯，如果走不出來，一生平庸。",
                ["庚丑"] = "十二月庚金，通根，喜見財官，木火相生，暖土生金，再走東南運，運順身體好，財運也不錯，火過旺，必見災。",
                // === 辛金 ===
                ["辛寅"] = "正月辛金，再有財官印比的幫扶，就會見官得官，見財得財。不喜見庚金，則有爭輝狀態，明爭暗鬥難免。",
                ["辛卯"] = "二月辛金，喜見丁火，坐支一片為土氣就更漂亮了，會出現奇才之人，但不能見比劫，見了比劫必然坎坷，一生小人比較多。",
                ["辛辰"] = "三月辛金，最喜見財，見殺，不能官殺混雜，不利仕途，易招官災。最喜獨財獨殺，必有名聲，財多了見財生災。土多也不行，雖然走出火土運比較好，但土多金埋，把金的光輝給埋沒了，發揮不來。",
                ["辛巳"] = "四月辛金，受制於月令，發揮光輝的時候，見丙火比較好，不能見丁火，丁火一見必化為水，見食傷不利前途，女命感情不順不利婚姻，行運能制住食傷則順利。",
                ["辛午"] = "五月辛金，七殺當令，有根且深的話必出官貴，特別是公檢法司方面，無根或根弱，再見官殺，則易生災，易牢獄之災。",
                ["辛未"] = "六月辛金，喜七殺配印，出文化人，但印不能太多，多了土多埋金，大多主牢獄之災、官司口舌，大運走出財官一氣，必是出名之時，七殺配印，根基雄厚的話，人才兩得。",
                ["辛申"] = "七月辛金，喜水，喜水來淘，不可見土，見土渾濁，淘不了金，金清水白財可以淘金，見了酉金強根，會有爭名爭利的現象，見丙火齊輝，見丁火為出奇觀，見水是比較好的。",
                ["辛酉"] = "八月辛金，自強喜丁火，丁火旺，子貴妻賢，財官雙得，運行水地，必發貴。這樣八字多生貴子，而且仕途比較好。",
                ["辛戌"] = "九月辛金，喜七殺配印，多出武職，福壽雙全，不怕障礙，運氣比較平穩，順勢安康，忌見巳火，運行巳火運，出意外奇災。",
                ["辛亥"] = "十月辛金，最喜丙火太陽，十月季節不怕官殺混雜，因為特寒特冷，辛金暗淡無光，喜火增輝。走南方運佳，走水運，用神受克不吉。",
                ["辛子"] = "十一月辛金，沒有丙丁火，運勢平平，運走南方也不會太大發展，比平常好一些，後天走出來的就不如先天了，運過去了也是廢。",
                ["辛丑"] = "十二月辛金，喜見火土，忌見金水，沒有火土，而且水成勢，壽不高，財不好，貧困之人，有了火土，大運流年一引動必出奇才。",
                // === 壬水 ===
                ["壬寅"] = "正月壬水，食神旺，天寒地凍，走南方運比較好。巳午未運增，財運佳。也不怕官殺，見到七殺反而能掌權，但是必須要見金，特別是申金。",
                ["壬卯"] = "二月壬水，傷官當令，寒氣還有，見火運佳。原局見火更好，起調候作用，大運走用神都是好的。",
                ["壬辰"] = "三月壬水，水土庫，有強根，透出甲乙木，食傷制殺，但是地支不能成木氣一片，就制太重，喜財喜印，原局中有財有印大運不忌。什麼運都比較順通，如果沒有金印，只有財，財殺一起就不順了。",
                ["壬巳"] = "四月壬水，偏財旺，行旺運為佳，從財更好，則忌西北運，為破局。",
                ["壬午"] = "五月壬水，火炎，水必須要有金生，喜比肩幫身，印助身都是上等格局。",
                ["壬未"] = "六月壬水，燥土見水吸干，見金生水，地支見申金就不怕。命局中沒有金水，就看能不能從了。",
                ["壬申"] = "七月壬水，財官易得，運行南方更強，但不喜走到絕地，大運走寅卯絕地，金在絕水在病必有一災。",
                ["壬酉"] = "八月壬水，正印當令，喜官殺，七殺配印好，地支怕見卯，卯木沖根不好，見辰也不好，辰酉合把根合跑了。",
                ["壬戌"] = "九月壬水，喜見金，財多殺多，官殺混雜就不好了，最忌諱身弱，財清富官清貴，無財無殺即使走到財鄉也不太好。",
                ["壬亥"] = "十月壬水，屬於見祿，喜走旺運，南方運雖然美好，運氣比較順當，但克妻財，克父。",
                ["壬子"] = "十一月壬水，喜財殺，無財無殺啥也不是，屬於流浪在外，無財無殺即使走到財鄉也不太好。",
                ["壬丑"] = "十二月壬水，喜見火財，財旺身旺必發富一方，走官發貴，忌金水增寒，而成冰，運氣也不通。",
                // === 癸水 ===
                ["癸寅"] = "正月癸水天寒，和地下河泥凍在一起，這種人抗壓力比較強，喜火暖身，大運走出金，金生水身旺，再見丙丁火必發無疑，原局中喜見印比，再見丙丁火為佳，忌見水，金也不成，金生水寒，增寒。也喜土，大運走出金喜歡，原局中不要見金。",
                ["癸卯"] = "二月癸水，氣開始往上竄了。地氣帶著水氣往上蒸發，有開拓精神，思想超前，不喜見官殺蓋頭，喜見丙丁火。",
                ["癸辰"] = "三月癸水，地下水已經通開了，比較聰明，辰為水庫，雜氣，四季庫挺好，什麼都能收納，可收可放，收放自如，身弱，這個月的癸水不怕弱，因為有個辰土的雜氣庫，弱可經霜，但不可見己土七殺，喜正官合，喜丙丁火，所以三月份的癸水是比較順暢的。",
                ["癸巳"] = "四月癸水，陰水地下水，不喜見陽光，見陽光就曬乾了，需要生，需要幫扶，如果無一點生扶，從了更好。",
                ["癸午"] = "五月癸水，火最旺，比較燥，喜印比，根深能盛財，如果弱就從財為佳，不強而弱或者假從，一生命運肯定是泥濘的，而且壽不長，所以五月一是從，要麼強，否則命不強壽不長。",
                ["癸未"] = "六月癸水，七殺當令，燥土一片，喜見印星，忌財多破印，印旺自身強，再走食傷運，這叫福祿雙豐，忌火土之鄉，會因財而傷。",
                ["癸申"] = "七月癸水，印當令，坐根強，見財官殺都可以，通根水木，必出財貴，水氣木氣都可以，三合局之類為通氣。",
                ["癸酉"] = "八月癸水，偏印身旺坐強，更喜丙丁火財鄉，忌印多梟食，財多也不良。印多貪生，人比較懶惰，財多則選擇性多了，癸水陰水多疑且有智慧，今天想干這，明天想干那，比較反復，這就是財多不為良。",
                ["癸戌"] = "九月癸水，正官當令，雜氣生髮，柱中必須要有印星，運行財鄉必富，運行官則貴，柱中喜見印比。喜生扶拱合，因九月的水正是乾燥之時，喜生扶。",
                ["癸亥"] = "十月癸水，劫財當令，喜官殺制比劫，否則一生坎坷，克財克父，生活艱難，象陷入泥潭一樣。忌北方運，見財官殺都比較好。",
                ["癸子"] = "十一月癸水，祿氣當權，喜見財運，柱中喜丙丁火為奇，忌西北方運，走南方火比較好。",
                ["癸丑"] = "十二月癸水，雜氣，喜見財見印，無財無印不為吉，南方火運最美。但臘月水已經成冰了，根土地已經和在一起了，無財無印不吉也沒大凶事，這樣人比較平穩，比較平淡。",
            };

            var sb2 = new System.Text.StringBuilder();
            if (stemBase.TryGetValue(dStem, out var sd) && !string.IsNullOrEmpty(sd))
                sb2.AppendLine(sd);
            if (monthDesc.TryGetValue(dStem + mBranch, out var md) && !string.IsNullOrEmpty(md))
                sb2.AppendLine(md);
            return sb2.ToString().Trim();
        }

        private static string LfPatternDesc(string pattern) => pattern switch
        {
            "正官格" => "命主天生具官貴氣質，一生有制度保障，適合公職管理，貴人多助。",
            "七殺格" => "命主魄力十足，敢冒險衝刺，適合軍警競爭行業，需注意制殺方能建功。",
            "正印格" => "命主學識深厚，貴人多助，一生靠學識名聲立足，財運平穩。",
            "偏印格" => "命主思想特殊，多才多藝，孤僻獨立，適合宗教、藝術、特殊技能方向。",
            "正財格" => "命主務實勤勞，財富穩定，重物質生活，一生衣食不缺。",
            "偏財格" => "命主豪爽善交際，善於冒險理財，財來財去，貿易業務見長。",
            "食神格" => "命主隨和享受，具藝術才能，適合餐飲、技術、創意行業。",
            "傷官格" => "命主聰明叛逆，才華橫溢，適合技藝創意自由業，婚姻感情需謹慎。",
            "建祿格" => "命主自強不息，靠自身努力打拼，財富靠雙手掙來，不喜依賴他人。",
            "月刃格" => "命主個性剛強，競爭意識強，財路需防劫財耗損，合夥宜謹慎。",
            "從強格" => "命主以從旺勢為吉，順從主流大方向，不宜逆勢而行。",
            "從殺格" => "命主日主極弱、七殺極旺，一生宜順從主流強勢，借力使力，忌逆勢抵抗。宜從事競爭型行業，但防財庫不穩。",
            "從財格" => "命主日主極弱、財星極旺，一生重財重物質，善於理財聚財，忌比劫印星破格。",
            "從兒格" => "命主日主極弱、食傷極旺，一生才藝豐沛，表達力強，宜創意技藝行業，忌印星梟奪。",
            "從旺格" => "命主極強，一生宜自主掌控，忌受人管束，順其旺勢大展。",
            "曲直格" => "命主木氣純粹，性格仁慈溫和，具人文藝術涵養，一生宜順木性發展。",
            "炎上格" => "命主火氣純粹，熱情積極，才華外顯，一生光芒四射，忌水剋而滅。",
            "稼穡格" => "命主土氣純粹，性格敦厚踏實，重情重義，一生宜從事土地農業相關。",
            "從革格" => "命主金氣純粹，性格剛毅果斷，有魄力，宜從事金融法律軍政。",
            "潤下格" => "命主水氣純粹，智慧流通，善變通達，宜從事流通貿易資訊行業。",
            _ => "命主格局中正，宜均衡發展，隨機應變。"
        };

        private static string LfPersonalityDesc(string dmElem, string pattern, double bodyPct)
        {
            var sb = new System.Text.StringBuilder();

            // 先天個性（日主五行）
            string elemDesc = dmElem switch
            {
                "木" => "木命之人如春木向陽，天生有理想有目標，重視道義，對朋友真誠，對自己有一定要求。思維清晰有條理，行事喜歡按計劃進行，一旦下定決心便難以更改。在人際方面偶因堅持原則而與人摩擦，但正因如此周遭的人都視您為可信賴的對象。有向上的拚勁，凡事不服輸，能在逆境中仍保持前進的意志。",
                "火" => "火命之人熱情外向，待人真誠，感情豐沛，善於表達，在人群中容易成為焦點。您直率坦白，不喜歡藏話，對自己認定的事情會全力投入。重視禮儀與場面，給人陽光開朗的印象；唯性子較急，有時說話直接讓人難以接受，情緒起伏明顯，高興時熱情如火，不順心時情緒也顯而易見。",
                "土" => "土命之人如厚土載物，穩重踏實，誠信待人，是旁人最信賴的朋友與夥伴。您處事沉穩不浮躁，能夠默默耕耘、長期堅持，凡事務實不好高騖遠。對於新環境、新事物的接受需要時間，但一旦熟悉了便能充分發揮實力。您對親近的人照顧有加，責任感強，是家人與朋友心中的靠山。",
                "金" => "金命之人如磨礪寶劍，個性剛直果決，重義氣，說話算話，言出必行。您有明確的是非觀，對自己要求甚嚴，品味講究，凡事追求品質而非數量。在人際關係上較為選擇性，一旦認定的朋友便全力相助，但也因此朋友圈相對固定。遇到不平之事容易義憤填膺，敢於直言，是非常有原則感的人。",
                "水" => "水命之人如流水靈活，天生聰明敏銳，思路廣博，善於學習，記憶力好。您能快速理解複雜事物，善於察言觀色，在人際往來中靈活圓融。唯思慮過多，容易想太多、擔心太多，有時反而錯失機會；情緒上偏向內斂，不輕易表達真實感受，外表看似平靜，內心波瀾起伏，需注意不要讓過多的憂慮影響行動力。",
                _ => ""
            };
            sb.AppendLine(elemDesc);
            sb.AppendLine();

            // 強弱特質
            string strDesc = bodyPct >= 65
                ? "日主旺強，您有強烈的自我意識與主導欲，凡事有主見，不喜受人擺佈，行動力強、意志堅定是您最大的優勢。需注意的是有時過於自信，容易聽不進他人意見；在做重大決定前多聆聽、多思考，可避免因固執而錯失良機。"
                : bodyPct <= 40
                ? "日主偏弱，個性較謙遜，懂得借力使力，貴人緣相對較好。在有組織有靠山的環境中往往能發揮更大潛力。需注意的是容易受外界影響情緒，面對壓力時偶有退縮傾向——培養內心的定力與自信，是您一生最重要的人生課題，一旦建立起這份自信，命局的優勢將充分展現。"
                : "日主強弱適中，個性均衡，既有主見又能聆聽，適應力強，能隨機應變，是處世比較從容的類型。在人際關係中圓融有彈性，少製造不必要的摩擦，能在各種環境中找到自己的位置。";
            sb.AppendLine(strDesc);
            sb.AppendLine();

            // 格局個性
            string patDesc = pattern switch
            {
                "正官格" => "格局為正官格，您有天生的規矩守法意識，重視名譽與形象，在有制度的環境中如魚得水。您對自己有高標準，管理能力突出，適合在機關或企業中擔任有名有份的職位。一生在乎別人如何評價自己，名聲對您而言比金錢更重要，也因此一生行事謹慎，不輕易踰越紅線。",
                "七殺格" => "格局為七殺格，您天生有魄力，敢衝敢拚，不畏壓力，壓力對您而言不是阻礙而是動力。在競爭激烈的環境中往往能脫穎而出，具備開創事業的基因。需注意情緒管理與人際溝通，避免行事太急或說話太衝，替自己製造不必要的阻力。",
                "食神格" => "格局為食神格，您天生愛享受生活，對飲食、藝術、美學有獨特品味，個性隨和不計較，是朋友眼中的開心果。您有藝術天份與創作力，適合在有創意空間的環境中工作，財運方面往往在享受生活的過程中自然積累，不需刻意強求。",
                "傷官格" => "格局為傷官格，您頭腦靈活，才華橫溢，思維不受常規束縛，有強烈的創新意識。對制度與權威天生有挑戰精神，適合從事需要創意或技術突破的工作。表達能力強，但有時說話過於直接犀利，容易得罪上司或長輩，建議在職場中適度收斂鋒芒，擇機而為。",
                "正財格" => "格局為正財格，您腳踏實地，務實重信，凡事一步一腳印，靠努力換取回報，不求僥倖。在理財方面有天份，懂得量入為出，長期積累財富。誠信待人，在商業往來中信譽良好，感情方面以穩定可靠為首要考量。",
                "偏財格" => "格局為偏財格，您人緣廣、交際廣，異性緣佳，善於與各種人打交道。財運偏向偏財，有機會透過業務、貿易、投資等方式獲取額外財富，但偏財來去無常，需注意守財。您慷慨大方、重情義，一生結交的人脈往往是最大的隱形資產。",
                "正印格" => "格局為正印格，您天生愛學習，求知欲旺盛，思維深沉，善於研究分析，有書卷氣，重視文化與學識。在知識型、文化型的環境中最能發揮長才，得長輩緣、貴人緣，往往能在關鍵時刻獲得指引。感情上略顯保守，不善主動表達，但情感深厚，一旦認定便忠誠可靠。",
                "偏印格" => "格局為偏印格，您有獨特的思維方式，往往與眾不同，對神秘學、哲學、藝術有天然的興趣，獨立性強。有超強的直覺力，在某些特殊領域往往有超出常人的洞察。建議找到一個能讓您深入鑽研的專業領域，在其中全情投入，方能發揮最大潛能。",
                _ => $"格局為{pattern}，宜依用神方向調整行事風格，以符合命局最佳發展路徑。"
            };
            sb.AppendLine(patDesc);

            return sb.ToString().TrimEnd();
        }

        private static string LfCareerDesc(string pattern) => pattern switch
        {
            "正官格" => "公職、管理、制度性工作",
            "七殺格" => "軍警、競爭性行業、創業",
            "食神格" => "餐飲、藝術、技術性工作",
            "傷官格" => "技藝、創意、自由業",
            "正財格" => "商業、財務、穩定收入行業",
            "偏財格" => "貿易、業務、投機性財富行業",
            "正印格" => "學術、教育、文職",
            "偏印格" => "宗教、藝術、特殊技能",
            _ => "多元發展方向"
        };

        private static string LfWealthDesc(double caiPct, double bodyPct, double biPct)
        {
            string baseDesc;
            if (caiPct >= 25 && bodyPct >= 60)
                baseDesc = "財多身強，財富根基厚實，一生衣食無憂，有能力取財也有能力守財。適合積極理財投資，事業越做越大。";
            else if (caiPct >= 25 && bodyPct < 40)
                baseDesc = "財多身弱，財路雖廣但體力財力難以承載，容易辛苦奔波而財散難留。最忌貪大求全，宜量力而為，先求守財再求進財，節流比開源更重要。";
            else if (biPct >= 30 && caiPct < 15)
                baseDesc = "比劫奪財，命中財星偏弱而同儕競爭力強，合夥容易出現利益糾紛，財路上宜獨立自主，慎防他人借財不還或感情破財。";
            else if (caiPct < 15 && bodyPct >= 60)
                baseDesc = "財星偏弱，日主強旺但財源有限，守成有餘、開源不足，財富積累靠長期耕耘而非一夕暴富。不宜冒進投資，穩健理財才是正道。";
            else
                baseDesc = "財富屬於中等水平，靠自身努力逐步積累，財運與個人努力程度高度相關。在喜用大運期間積極進取，可有所突破；凶運期間守住已有的財富即是勝利。";
            string advice = bodyPct < 50
                ? "建議：避免單打獨鬥，藉助貴人或合夥之力放大財運，善用人脈資源。"
                : "建議：財運好壞與行運密切相關，喜用運積極開展，忌神運低調保守。";
            return baseDesc + "\n" + advice;
        }

        private static string LfMarriageDesc(double spousePct, string dBranch, string dStem, string dmElem, int gender, string[] branches)
        {
            var sb = new System.Text.StringBuilder();
            bool anyChong = branches.Where(b => b != dBranch).Any(b => LfChong.Contains(dBranch + b));
            string spouseElem = gender == 1
                ? LfElemOvercome.GetValueOrDefault(dmElem, "")
                : LfElemOvercomeBy.GetValueOrDefault(dmElem, "");

            // 配偶特質
            string spouseChar = spouseElem switch
            {
                "木" => "配偶五行屬木，對方個性有上進心，重視道義，溫和但有原則，外表親和、內心堅定，是有理想的類型。",
                "火" => "配偶五行屬火，對方個性熱情爽朗，表達能力強，重視感情互動，有活力有魄力，婚後生活較為多彩多姿。",
                "土" => "配偶五行屬土，對方個性穩重踏實，顧家負責，是居家型的好伴侶，婚後生活穩定，重視家庭責任與實際承諾。",
                "金" => "配偶五行屬金，對方個性剛直果決，重情義，品味不俗，有主見，感情上較為直接，不善拐彎抹角，是說到做到的類型。",
                "水" => "配偶五行屬水，對方個性聰明靈活，善解人意，情感細膩，善於照顧另一半的情緒；但有時想法較多、情緒起伏也較難捉摸。",
                _ => ""
            };
            if (!string.IsNullOrEmpty(spouseChar)) { sb.AppendLine(spouseChar); sb.AppendLine(); }

            // 婚緣強弱
            string strengthDesc = spousePct >= 20
                ? $"命局婚緣不弱（婚星占 {spousePct:F0}%），感情路上異性緣較好，有機會遇到合適的對象，感情生活較有依靠。婚後只要雙方相互珍惜，婚姻關係有一定穩定性。"
                : $"命局婚緣偏薄（婚星占 {spousePct:F0}%），感情路上波折略多，緣分到來需要好好珍惜，不宜輕易放棄或優柔寡斷。婚後需要雙方更多包容與耐心，才能維持感情長久穩定。";
            sb.AppendLine(strengthDesc);
            sb.AppendLine();

            // 日支夫妻宮
            if (anyChong)
                sb.AppendLine("日支（夫妻宮）逢沖，婚姻感情容易出現分離、爭執或觀念不合的情形。這並非不能婚，而是需要雙方付出更多的理解與包容，尤其在大運或流年沖動日支的年份，婚姻關係容易出現波動，宜提前溝通化解。");
            else
                sb.AppendLine("日支（夫妻宮）較為平穩，婚姻宮位無大沖刑，夫妻感情有一定的穩定基礎，只要雙方用心維繫，感情可以細水長流。");
            sb.AppendLine();

            // 感情建議
            string advice = gender == 1
                ? "感情建議：行財星旺運（喜用年份）是感情婚姻機會最大的時機。婚後宜以體貼包容代替強勢主導，尊重另一半的意見，是婚姻長久的關鍵。"
                : "感情建議：行官星旺運（喜用年份）是感情婚姻機會最大的時機。婚後在生活細節上用心，避免過於強勢或疑心過重，給雙方空間，婚姻才能穩固。";
            sb.AppendLine(advice);

            string dBranchMainElem = LfBranchHiddenRatio.TryGetValue(dBranch, out var bhe) && bhe.Count > 0 ? KbStemToElement(bhe[0].stem) : "";
            string dStemElem = KbStemToElement(dStem);
            if (gender == 1 && LfElemOvercome.GetValueOrDefault(dStemElem, "") == dBranchMainElem)
            {
                sb.AppendLine();
                sb.AppendLine("日干克日支，命主個性較強，婚後兩人個性差異明顯，建議主動退讓，避免以己之長攻彼之短，才能維持夫妻和諧。");
            }
            return sb.ToString().TrimEnd();
        }

        private static string LfHealthDesc(Dictionary<string, double> wuXing, string seaLabel)
        {
            var sb = new StringBuilder();
            var organMap = new Dictionary<string, string>
            {
                { "木","肝膽、眼睛、筋骨" }, { "火","心臟、血液、血壓" },
                { "土","脾胃、消化系統、肌肉" }, { "金","肺、呼吸系統、皮膚" }, { "水","腎臟、膀胱、骨骼" }
            };
            foreach (var (elem, organ) in organMap)
            {
                double pct = wuXing.GetValueOrDefault(elem, 0);
                if (pct >= 35) sb.AppendLine($"  {elem}旺（{pct:F0}%）：{organ}易亢進，注意過旺之症。");
                else if (pct <= 8) sb.AppendLine($"  {elem}弱（{pct:F0}%）：{organ}易虛損，注意補養。");
            }
            if (seaLabel == "炎熱") sb.AppendLine("  體性炎熱：易有心血管、高血壓問題，年老後尤需注意防暑。");
            if (seaLabel == "寒凍") sb.AppendLine("  體性寒凍：易有腎虛、關節退化、循環系統問題，注意保暖。");
            if (sb.Length == 0) sb.AppendLine("  五行分布均衡，體質較為平和，注意日常保健即可。");
            sb.Append("  （注：以保健提醒為主，不作疾病診斷）");
            return sb.ToString().TrimEnd();
        }

        private static string LfLuckDesc(int score, string level) => level switch
        {
            "大吉運" => "黃金發展期，喜用五行大旺，事業財運大展，宜積極進取，把握機遇擴展版圖。",
            "中吉運" => "喜用五行得力，運勢上揚，有貴人助力，宜積極行動，可有所成就。",
            "平吉運" => "喜用五行略佔優勢，平中帶吉，宜穩健進取，持續耕耘自有收獲。",
            "平運"   => "喜忌五行相當，平穩守成，無大得失，宜穩健行事，蓄勢待發，厚積薄發。",
            "中凶運" => "忌神五行偏旺，喜用受制，宜保守行事，低調蓄積，謹防輕率決策造成損失。",
            "大凶運" => "忌神五行大旺，喜用全無，事業財運嚴重受阻，宜保守守成，謹防重大損失。",
            _ => ""
        };

        private static string LfKeyYears(
            List<(string stem, string branch, string liuShen, int startAge, int endAge, int score, string level)> scored,
            int birthYear, string yongElem, string jiElem)
        {
            var sb = new StringBuilder();
            var good = scored.Where(c => c.score >= 70).OrderByDescending(c => c.score).Take(2).ToList();
            var bad  = scored.Where(c => c.score < 50).OrderBy(c => c.score).Take(3).ToList();
            if (good.Count > 0)
            {
                sb.AppendLine("重點吉運建議把握：");
                foreach (var c in good)
                    sb.AppendLine($"  {birthYear + c.startAge}-{birthYear + c.endAge} 年（{c.startAge}-{c.endAge} 歲）{c.stem}{c.branch}大運 {c.score} 分：宜積極擴展事業、財運、感情。");
            }
            if (bad.Count > 0)
            {
                sb.AppendLine("重點凶運需謹慎：");
                foreach (var c in bad)
                    sb.AppendLine($"  {birthYear + c.startAge}-{birthYear + c.endAge} 年（{c.startAge}-{c.endAge} 歲）{c.stem}{c.branch}大運 {c.score} 分：此期宜保守守成，防財損、健康、人際波折。");
            }
            if (sb.Length == 0) sb.Append("整體大運尚稱平穩，依各大運評分適度調整行事策略。");
            return sb.ToString().TrimEnd();
        }

        private static string LfPeriodDesc(double avg) => avg switch
        {
            >= 90 => "運勢極旺，天時地利人和，此段為人生黃金期，宜大膽開創積極進取。",
            >= 85 => "運勢極強，貴人扶持，機遇頻現，宜主動出擊把握良機。",
            >= 80 => "運勢旺盛，諸事較為順遂，宜積極佈局擴大格局。",
            >= 75 => "運勢良好，整體走勢向上，偶有小阻，穩中可求大進。",
            >= 70 => "運勢偏吉，大方向順遂，細節需留意，宜穩步前行。",
            >= 65 => "運勢中上，起伏適中，把握時機仍可有所成就。",
            >= 60 => "運勢平穩，平中有吉，宜守成待機，蓄積實力。",
            >= 55 => "運勢平平，有吉有凶，宜謹慎行事，避免冒進。",
            >= 50 => "運勢中平，考驗與機遇並存，心態穩定是關鍵。",
            >= 45 => "運勢偏弱，考驗較多，宜低調行事，以守為主。",
            >= 40 => "運勢偏凶，起伏較大，宜謹慎保守，化解阻力。",
            >= 35 => "運勢較差，多遇波折，宜靜待轉機，切忌輕舉妄動。",
            _ => "運勢艱難，諸事考驗多，宜韜光養晦，厚積薄發。"
        };

        // Age-band specific advice for Ch.10, based on life-stage topics + 十神 in cycles
        private static string LfAgeBandAdvice(
            int ageStart, double avg, int gender, string yongElem, string jiElem, string dStem,
            List<(string stem, string branch, string liuShen, int startAge, int endAge, int score, string level)> cycles)
        {
            bool hasYin  = cycles.Any(c => { var ss = LfStemShiShen(c.stem, dStem); return ss == "正印" || ss == "偏印"; });
            bool hasGuan = cycles.Any(c => { var ss = LfStemShiShen(c.stem, dStem); return ss == "正官" || ss == "七殺"; });
            bool hasCai  = cycles.Any(c => { var ss = LfStemShiShen(c.stem, dStem); return ss == "正財" || ss == "偏財"; });
            bool hasJiInStem = cycles.Any(c => LfElemStems(jiElem).Contains(c.stem));

            if (ageStart == 0) // 0-13 童年學藝期
            {
                if (avg >= 70)
                    return "此期身體強健，成長順遂，才藝學習事半功倍，父母助力足。宜積極培養一技之長，奠定日後發展基礎。";
                if (avg >= 55)
                    return "身體大致平穩，偶有小恙，才藝培育宜循序漸進，注意交通安全與意外傷害，父母多加留心。";
                return "此期身體較弱，易有意外或疾病，父母宜多留意安全健康，才藝學習量力而為，莫強求。";
            }

            if (ageStart == 14) // 14-18 青少成長期
            {
                string studyPart = hasYin ? "印星入運，升學考試有利，學業表現亮眼，"
                                  : hasJiInStem ? "忌神干擾，學業壓力稍大，考試需加倍努力，"
                                  : "學業表現中規中矩，";
                if (avg >= 70)
                    return $"{studyPart}此期才藝發展順遂，感情初萌可自然發展，有出國讀書或遊學機遇，宜把握。";
                if (avg >= 55)
                    return $"{studyPart}感情宜晚，以學業才藝為重，出國機遇視大運配合而定，保持穩定為宜。";
                return $"{studyPart}宜低調努力，感情暫緩，避免分心，先穩固學業基礎，出國計劃謹慎評估。";
            }

            if (ageStart == 19) // 19-25 青年立志期
            {
                if (avg >= 65 && hasYin)
                    return "印星扶助，此期最適合繼續深造，碩博可期，學術成就佳；若選擇就業，學習能力強，晉升機遇多。";
                if (avg >= 65 && hasGuan)
                    return "官殺入運，投入職場體系最為有利，循規蹈矩可獲晉升；也可考慮深造以強化競爭力。";
                if (avg >= 65)
                    return "整體運勢強，可評估深造或就業，順應自身格局選擇，此期奠定的基礎將影響未來走向。";
                if (avg >= 50)
                    return hasYin ? "雖有印星之利，但運勢稍弱，宜穩步深造或紮實學習技能，切勿好高騖遠，厚積薄發。"
                                  : "運勢平平，宜先就業積累一技之長，打好根基，待大運轉強再圖深造或晉升。";
                return "此期起步考驗多，建議先就業學習實際技能，打好根基，待大運轉強再謀深造或升遷。";
            }

            if (ageStart == 26) // 26-35 成家立業期
            {
                string spouseStar = gender == 1 ? "財星（妻星）" : "官星（夫星）";
                bool hasSpouseStar = gender == 1 ? hasCai : hasGuan;
                string marriagePart = hasSpouseStar && avg >= 60
                    ? $"{spouseStar}入運，姻緣水到渠成，此期婚配順利，宜把握感情適時論嫁娶。"
                    : avg >= 65 ? $"感情發展中，{spouseStar}配合尚可，婚配時機需觀察流年共同判斷。"
                    : $"此期{spouseStar}受制，感情易有波折，婚配宜謹慎評估，不急於一時。";
                string careerPart = avg >= 65 ? "事業同步奠基，財富逐漸積累，置產時機可待有利流年。"
                                  : avg >= 50 ? "事業仍在打拼，財務量力而為，置產計劃宜謹慎評估。"
                                  : "事業遭逢考驗，財務保守為宜，婚配與置產均需量力而行。";
                return $"{marriagePart} {careerPart}";
            }

            if (ageStart == 36) // 36-50 壯年拼搏期
            {
                if (avg >= 75)
                    return "壯年黃金期，事業進入高峰，財運亨通，宜積極拓展版圖，購屋置產時機佳；注意維繫家庭關係，子女教育同步用心。";
                if (avg >= 60)
                    return "事業穩健推進，財運中平，守成為主，量力投資切勿冒進；關注家庭子女教育，定期健康檢查，注意出行安全。";
                return "此期事業財運考驗較多，宜保守理財、鞏固現有根基，注意身體健康與意外風險，家庭關係多加維繫，以穩為要。";
            }

            // 51+ 晚年守成期
            if (avg >= 70)
                return "晚運旺盛，田宅家業穩固，身體康健，宜保持良好生活習慣，享受家庭天倫，適度養生，晚年安樂。";
            if (avg >= 55)
                return "晚運平穩，家庭生活安定，宜定期健康檢查，注意慢性病預防；財務量力而行，安養天年為要。";
            return "晚運考驗較多，身體宜多保養，財務保守，避免高風險投資，家人陪伴與照護至關重要，晚年以健康為首。";
        }

        private static string LfWealthLevel(double caiPct, double bodyPct)
        {
            if (caiPct >= 25 && bodyPct >= 60) return "豐厚";
            if (caiPct >= 15) return "中等";
            return "平淡";
        }

        private static string LfFameLevel(string pattern) => pattern switch
        {
            "正官格" or "七殺格" => "顯達", "正印格" or "偏財格" or "食神格" => "小成", _ => "平凡"
        };

        private static string LfElemDir(string elem) => elem switch
        { "木"=>"東方","火"=>"南方","土"=>"中央","金"=>"西方","水"=>"北方",_=>"" };

        private static string LfElemColor(string elem) => elem switch
        { "木"=>"綠色、青色","火"=>"紅色、橙色","土"=>"黃色、棕色","金"=>"白色、金色","水"=>"黑色、深藍色",_=>"" };

        private static string LfElemCareer(string elem) => elem switch
        {
            "木" => "教育、文學、醫療、植物相關行業",
            "火" => "表演、能源、食品、熱門行業",
            "土" => "房地產、農業、建設、倉儲行業",
            "金" => "金融、機械、法律、IT 行業",
            "水" => "貿易、流通、餐飲、娛樂行業",
            _ => ""
        };

        private static string LfYearDesc(int score) => score >= 70
            ? "早年家境良好，祖業有助，父母有力，出身較佳。"
            : score >= 50 ? "早年家境平常，父母中等，靠自身努力為主。"
            : "早年家境較艱苦，父母緣薄，少有祖業蔭助，自力更生。";

        private static string LfMonthDesc(int score) => score >= 70
            ? "父母感情和睦，兄弟姐妹有情誼，青年期（16-31 歲）發展順利。"
            : score >= 50 ? "父母關係尚可，兄弟各有際遇，青年期有波折但能克服。"
            : "父母緣分較薄，兄弟感情不睦，青年期多有挫折，需自強。";

        private static string LfDayDesc(int score, int gender) => gender == 1
            ? (score >= 70 ? "妻星有力，感情豐富，婚姻和諧，妻子賢良有助力。"
               : score >= 50 ? "妻緣尚可，婚姻有小波折，夫妻需多包容溝通。"
               : "妻緣偏弱，感情路多曲折，婚姻需謹慎選擇。")
            : (score >= 70 ? "夫星有力，夫緣深厚，丈夫能幹有助力，婚姻和諧。"
               : score >= 50 ? "夫緣尚可，婚姻有小磨擦，需多溝通包容。"
               : "夫緣偏弱，感情路多波折，婚姻宜謹慎。");

        private static string LfHourDesc(int score) => score >= 70
            ? "子女孝順賢能，晚年（48-65 歲）有子女助力，老年生活安康。"
            : score >= 50 ? "子女緣中等，晚年尚有依靠，但需自己多積蓄。"
            : "子女緣薄，晚年多靠自身，需提前做好晚年規劃。";

        // === Appendix E Rule Engine ===

        private static string LfApplyRules(
            string yStem, string yBranch, string mStem, string mBranch,
            string dStem, string dBranch, string hStem, string hBranch,
            string yStemSS, string mStemSS, string hStemSS,
            string yBranchSS, string mBranchSS, string dBranchSS, string hBranchSS,
            string dmElem, Dictionary<string, double> wuXing, int gender,
            string pattern, double bodyPct, string[] branches)
        {
            var sb = new StringBuilder();

            // E-1 Father rules
            if ((yStemSS is "比肩" or "劫財") && (yBranchSS is "比肩" or "劫財"))
                sb.AppendLine("  【F01 父緣】年柱干支皆比劫，比劫奪財，父緣薄或早年少父蔭。");
            if (yStemSS == "偏財" && !branches.Where(b => b != yBranch).Any(b => LfChong.Contains(yBranch + b)))
                sb.AppendLine("  【F04 父緣】年干偏財透出且根穩，父親有財有能力，早年家境良好。");

            // E-1 Mother rules
            bool hasZhengYin = yStemSS == "正印" || mStemSS == "正印" || hStemSS == "正印";
            bool mBranchChong = branches.Where(b => b != mBranch).Any(b => LfChong.Contains(mBranch + b));
            if (hasZhengYin && !mBranchChong)
                sb.AppendLine("  【M01 母緣】正印星透出有力，母親賢慧，早年得母助，與母感情深厚。");

            // E-2 Sibling rules
            double biPct = wuXing.GetValueOrDefault(dmElem, 0);
            if (biPct >= 30 && (mStemSS is "比肩" or "劫財") && (mBranchSS is "比肩" or "劫財"))
                sb.AppendLine("  【B01 兄弟】比劫旺且月柱皆比劫，兄弟多但財路易受影響，合夥需謹慎。");
            if (biPct < 8)
                sb.AppendLine("  【B04 兄弟】比劫偏弱，兄弟緣薄，孤立無援，少有兄弟助力。");
            if (LfChong.Contains(mBranch + dBranch))
                sb.AppendLine("  【B03 兄弟】月支沖日支，與兄弟分離之象，少年時期即各奔東西。");

            // E-3 Spouse rules
            if (branches.Any(b => b != dBranch && LfChong.Contains(dBranch + b)))
                sb.AppendLine("  【S01 配偶】日支逢沖，夫妻有分離之象，婚姻多波折，需多包容。");
            if (gender == 1)
            {
                string dStemElem = KbStemToElement(dStem);
                string dBrMainElem = LfBranchHiddenRatio.TryGetValue(dBranch, out var bhe) && bhe.Count > 0
                    ? KbStemToElement(bhe[0].stem) : "";
                if (LfElemOvercome.GetValueOrDefault(dStemElem, "") == dBrMainElem)
                    sb.AppendLine("  【S02 配偶】日干克日支（男命），有剋妻之象，妻身體偏弱或婚後多摩擦。");
                bool hasZhengCai = yStemSS == "正財" || mStemSS == "正財" || hStemSS == "正財";
                bool hasPianCai  = yStemSS == "偏財" || mStemSS == "偏財" || hStemSS == "偏財";
                if (hasZhengCai && hasPianCai)
                    sb.AppendLine("  【S05 配偶】正偏財均透干，二婚之象或感情複雜，需謹慎對待感情。");
            }
            else if (pattern == "傷官格")
                sb.AppendLine("  【S06 配偶】傷官格（女命），傷官見官，婚姻多波折，較晚婚或有離婚之象。");

            // E-4 Children rules
            if (LfChong.Contains(hBranch + dBranch))
                sb.AppendLine("  【C01/C02 子女】時支沖日支，子女與父母緣薄，晚年需自立自強。");
            string childElem = gender == 1 ? LfElemGen.GetValueOrDefault(dmElem, "") : LfElemOvercomeBy.GetValueOrDefault(dmElem, "");
            if (wuXing.GetValueOrDefault(childElem, 0) < 8)
                sb.AppendLine("  【C03 子女】子女星五行偏弱，子女緣薄，少子或晚得子。");

            // E-5 Career/Wealth rules
            double caiPct = wuXing.GetValueOrDefault(LfElemOvercome.GetValueOrDefault(dmElem, ""), 0);
            if (caiPct >= 25 && bodyPct >= 60)
                sb.AppendLine("  【W01 財運】財多身強，財富豐厚，一生衣食無憂。");
            else if (caiPct >= 25 && bodyPct < 40)
                sb.AppendLine("  【W02 財運】財多身弱，雖有財路但難守住，需強化理財觀念。");
            if (biPct >= 30 && caiPct < 15)
                sb.AppendLine("  【W03 財運】比劫奪財，財路受阻，合夥易損，獨立經營較佳。");

            // E-6 Health rules
            foreach (var xg in new[] { new[] { "寅","巳","申" }, new[] { "丑","戌","未" } })
                if (xg.All(b => branches.Contains(b)))
                    sb.AppendLine("  【H01 健康】三刑齊全，開刀手術之象，一生宜留意外傷或手術風險。");

            return sb.ToString().TrimEnd();
        }

        // === Dy (大運命書) Helper Methods ===

        // 輔星縮寫 → 全名對照
        private static readonly Dictionary<string, string> AuxStarExpand = new()
        {
            {"曲","文曲"},{"昌","文昌"},{"輔","左輔"},{"弼","右弼"},{"魁","天魁"},{"鉞","天鉞"},
            {"羊","擎羊"},{"陀","陀羅"},{"火","火星"},{"鈴","鈴星"}
        };

        // 取宮位內的六吉輔星與四剎星（無主星時的替代論斷依據）
        private static (List<string> goodStars, List<string> badStars) DyGetPalaceAuxiliaryStars(
            JsonElement palaces, string palaceName)
        {
            var good = new List<string>();
            var bad  = new List<string>();
            if (palaces.ValueKind != JsonValueKind.Array) return (good, bad);
            var sixGood = new HashSet<string> { "文昌","文曲","左輔","右弼","天魁","天鉞" };
            var fourBad = new HashSet<string> { "擎羊","陀羅","火星","鈴星" };
            foreach (var p in palaces.EnumerateArray())
            {
                string pname = p.TryGetProperty("palaceName", out var pn) ? pn.GetString() ?? "" :
                               p.TryGetProperty("name", out var n2) ? n2.GetString() ?? "" : "";
                if (!KbPalaceSame(pname, palaceName)) continue;
                foreach (var key in new[] { "secondaryStars", "goodStars", "badStars", "smallStars", "majorStars", "mainStars" })
                {
                    if (!p.TryGetProperty(key, out var arr) || arr.ValueKind != JsonValueKind.Array) continue;
                    foreach (var s in arr.EnumerateArray())
                    {
                        string sn = s.GetString() ?? "";
                        if (string.IsNullOrEmpty(sn)) continue;
                        // 縮寫展開 & 去除"星"後綴
                        string full = AuxStarExpand.TryGetValue(sn, out var exp) ? exp : sn.TrimEnd('星');
                        if (sixGood.Contains(full) && !good.Contains(full)) good.Add(full);
                        else if (fourBad.Contains(full) && !bad.Contains(full)) bad.Add(full);
                    }
                }
                break;
            }
            return (good, bad);
        }

        // 取指定地支的宮位名稱
        private static string KbGetPalaceByBranch(JsonElement palaces, string branch)
        {
            if (palaces.ValueKind != JsonValueKind.Array) return "";
            foreach (var p in palaces.EnumerateArray())
            {
                string raw = p.TryGetProperty("earthlyBranch", out var br) ? br.GetString() ?? "" : "";
                string b   = ExtractBranchChar(raw);
                if (b != branch) continue;
                string pn = p.TryGetProperty("palaceName", out var pnProp) ? pnProp.GetString() ?? "" :
                            p.TryGetProperty("name", out var n2) ? n2.GetString() ?? "" : "";
                return KbNormalizePalaceName(pn);
            }
            return "";
        }

        private static readonly string[] ZiweiBranchOrder = { "子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥" };

        // 計算大限財宮與事業宮（固定為 +4 / +8 地支偏移，無論順逆行）
        private static (string financePalace, string financeStars, string careerPalace, string careerStars)
            DyGetDecadeKeyPalaceStars(JsonElement palaces, string decadePalaceName)
        {
            string decadeBranch = KbGetPalaceBranch(palaces, decadePalaceName);
            if (string.IsNullOrEmpty(decadeBranch)) return ("","","","");
            int idx = Array.IndexOf(ZiweiBranchOrder, decadeBranch);
            if (idx < 0) return ("","","","");
            string financeBranch = ZiweiBranchOrder[(idx + 8) % 12];  // 財宮 = 逆推 -4 ≡ +8
            string careerBranch  = ZiweiBranchOrder[(idx + 4) % 12];  // 事業宮 = 逆推 -8 ≡ +4
            string financePalace = KbGetPalaceByBranch(palaces, financeBranch);
            string careerPalace  = KbGetPalaceByBranch(palaces, careerBranch);
            string financeStars  = string.IsNullOrEmpty(financePalace) ? "" : KbGetPalaceStars(palaces, financePalace);
            string careerStars   = string.IsNullOrEmpty(careerPalace)  ? "" : KbGetPalaceStars(palaces, careerPalace);
            // 取輔星補充（若主星為空）
            if (string.IsNullOrEmpty(financeStars))
            {
                var (fg, fb) = DyGetPalaceAuxiliaryStars(palaces, financePalace);
                financeStars = string.Join("、", fg.Concat(fb));
            }
            if (string.IsNullOrEmpty(careerStars))
            {
                var (cg, cb) = DyGetPalaceAuxiliaryStars(palaces, careerPalace);
                careerStars = string.Join("、", cg.Concat(cb));
            }
            return (financePalace, financeStars, careerPalace, careerStars);
        }

        // 六吉輔星與四剎星的宮位論斷注解
        private static string DyAuxStarNote(string star) => star switch
        {
            "文昌" => "才藝聰明，文書考試有利，名聲加分",
            "文曲" => "口才才藝出眾，學習力強，有異才",
            "左輔" => "得人緣貴人相助，輔佐力旺",
            "右弼" => "有助力，得異性緣，暗中有貴人",
            "天魁" => "日間貴人星，逢凶化吉，官場有利",
            "天鉞" => "夜間貴人星，逢凶化吉，得女性貴人",
            "擎羊" => "剛強衝動，橫發橫破，需防意外傷病與是非",
            "陀羅" => "延遲拖延，暗耗糾纏，凡事宜提早規劃",
            "火星" => "急躁衝動，爆發力強，需防火傷意外，亦可激發潛能",
            "鈴星" => "陰沉伏發，突發破耗，需防暗中損失",
            _      => ""
        };

        // 取得與八字大運年齡段重疊的所有紫微大限宮位（可能跨宮）
        // 回傳：(宮位名, 宮干, 大限起始歲, 大限結束歲) 依歲數排序
        private static List<(string palaceName, string palaceStem, int pStart, int pEnd)> DyGetOverlappingDecadePalaces(
            JsonElement palaces, int baziStart, int baziEnd)
        {
            var result = new List<(string, string, int, int)>();
            if (palaces.ValueKind != JsonValueKind.Array) return result;
            foreach (var p in palaces.EnumerateArray())
            {
                string range = p.TryGetProperty("decadeAgeRange", out var dr) ? dr.GetString() ?? "" : "";
                var parts = range.Split('-');
                if (parts.Length != 2 || !int.TryParse(parts[0], out int ps) || !int.TryParse(parts[1], out int pe)) continue;
                // 判斷是否與八字大運重疊
                if (ps > baziEnd || pe < baziStart) continue;
                string pname = p.TryGetProperty("palaceName", out var pn) ? pn.GetString() ?? "" :
                               p.TryGetProperty("name",       out var n2) ? n2.GetString() ?? "" : "";
                pname = KbNormalizePalaceName(pname);
                string stem = p.TryGetProperty("palaceStem", out var stemProp) ? stemProp.GetString() ?? "" : "";
                result.Add((pname, stem, ps, pe));
            }
            return result.OrderBy(x => x.Item3).ToList();
        }

        private static (string stem, string branch) DyGetYearStemBranch(int year)
        {
            string[] stems    = { "甲","乙","丙","丁","戊","己","庚","辛","壬","癸" };
            string[] branches = { "子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥" };
            int idx = (year - 4) % 60;
            if (idx < 0) idx += 60;
            return (stems[idx % 10], branches[idx % 12]);
        }

        private static int DyCalcFlowYearBaziScore(
            string flStem, string flBranch,
            string pattern, string yongShenElem, string fuYiElem, string jiShenElem,
            string dmElem, bool isBodyStrong, string tiaoHouElem,
            string season, string[] chartBranches, string[] chartStems)
        {
            var (goodElems, badElems) = LfGetPatternLuckElems(pattern, yongShenElem, fuYiElem, dmElem, isBodyStrong);
            if (!string.IsNullOrEmpty(tiaoHouElem) && !goodElems.Contains(tiaoHouElem) && !badElems.Contains(tiaoHouElem))
                goodElems = goodElems.Concat(new[] { tiaoHouElem }).ToArray();

            double score = 50.0;

            // 天干評分
            string flStemElem = KbStemToElement(flStem);
            if (!string.IsNullOrEmpty(flStemElem))
            {
                double mult = LfIsElemNeutralizedByChart(flStemElem, chartStems, chartBranches) ? 0.5 : 1.0;
                if (goodElems.Contains(flStemElem))     score += 18 * mult;
                else if (badElems.Contains(flStemElem)) score -= 18 * mult;
            }

            // 地支評分（依藏干比例）
            string flBranchMainElem = "";
            if (LfBranchHiddenRatio.TryGetValue(flBranch, out var flBH))
            {
                if (flBH.Count > 0) flBranchMainElem = KbStemToElement(flBH[0].stem);
                foreach (var (hstem, ratio) in flBH)
                {
                    string e = KbStemToElement(hstem);
                    if (string.IsNullOrEmpty(e)) continue;
                    double mult = LfIsElemNeutralizedByChart(e, chartStems, chartBranches) ? 0.5 : 1.0;
                    if (goodElems.Contains(e))     score += 18 * ratio * mult;
                    else if (badElems.Contains(e)) score -= 18 * ratio * mult;
                }
            }

            // 調候補助
            string tuneElem = season == "冬" ? "火" : season == "夏" ? "水" : "";
            if (!string.IsNullOrEmpty(tuneElem))
            {
                if (flStemElem == tuneElem && !goodElems.Contains(tuneElem))
                    score += badElems.Contains(tuneElem) ? 5 : 7;
                if (!string.IsNullOrEmpty(flBranchMainElem) && flBranchMainElem == tuneElem
                    && flBranchMainElem != flStemElem && !goodElems.Contains(tuneElem))
                    score += badElems.Contains(tuneElem) ? 3 : 4;
            }

            // 歲君沖/合/刑 微調
            double adj = 0;
            if (chartBranches.Any(b => LfChong.Contains(flBranch + b)))
            {
                // 歲君沖命局：若沖的是忌神地支 → 輕微加分；否則扣分
                bool hitsBadBranch = chartBranches.Any(b => LfChong.Contains(flBranch + b)
                    && LfBranchHiddenRatio.TryGetValue(b, out var bh) && bh.Count > 0
                    && badElems.Contains(KbStemToElement(bh[0].stem)));
                adj += hitsBadBranch ? 3 : -8;
            }
            if (LfHe.TryGetValue(flBranch, out var heInf) && chartBranches.Contains(heInf.partner))
                adj += goodElems.Contains(heInf.elem) ? 6 : (badElems.Contains(heInf.elem) ? -4 : 3);
            foreach (var xg in LfXing)
                if (xg.Contains(flBranch) && xg.Count(b => b != flBranch && chartBranches.Contains(b)) >= 1)
                { adj -= 5; break; }
            if (chartBranches.Any(b => LfHai.Contains(flBranch + b))) adj -= 3;
            if (chartBranches.Any(b => LfPo.Contains(flBranch + b))) adj -= 2;

            return (int)Math.Round(Math.Clamp(score + adj, 0, 100));
        }

        private static int DyCalcZiweiScore(string flStem, JsonElement palaces, string daiyunStem = "", int age = 0)
        {
            double score = 50.0;
            var keyPalaces  = new HashSet<string> { "命宮", "財帛宮", "官祿宮", "夫妻宮", "田宅宮" };
            var goodPalaces = new HashSet<string> { "福德宮", "父母宮" };
            string luPalace   = KbGetSiHuaPalace(flStem, "化祿", palaces);
            string quanPalace = KbGetSiHuaPalace(flStem, "化權", palaces);
            string kePalace   = KbGetSiHuaPalace(flStem, "化科", palaces);
            string jiPalace   = KbGetSiHuaPalace(flStem, "化忌", palaces);
            // 流年化祿（正向加分：年齡不合的宮位跳過，避免高齡者夫妻宮化祿虛增評分）
            if (!LfShouldSkipPalace(luPalace, age))
            {
                if (keyPalaces.Contains(luPalace)) score += 20;
                else if (goodPalaces.Contains(luPalace)) score += 12;
                else if (!string.IsNullOrEmpty(luPalace)) score += 6;
            }
            // 流年化權
            if (!LfShouldSkipPalace(quanPalace, age))
            {
                if (keyPalaces.Contains(quanPalace)) score += 12;
                else if (goodPalaces.Contains(quanPalace)) score += 7;
                else if (!string.IsNullOrEmpty(quanPalace)) score += 4;
            }
            // 流年化科
            if (!LfShouldSkipPalace(kePalace, age))
            {
                if (keyPalaces.Contains(kePalace)) score += 6;
                else if (!string.IsNullOrEmpty(kePalace)) score += 3;
            }
            // 流年化忌（化忌不過濾：任何年齡財庫/健康受損都有意義）
            if (keyPalaces.Contains(jiPalace)) score -= 20;
            else if (goodPalaces.Contains(jiPalace)) score -= 12;
            else if (!string.IsNullOrEmpty(jiPalace)) score -= 6;
            // 大限宮干四化疊加（大限結構性影響，力道約為流年一半）
            if (!string.IsNullOrEmpty(daiyunStem))
            {
                string dyLu   = KbGetSiHuaPalace(daiyunStem, "化祿", palaces);
                string dyQuan = KbGetSiHuaPalace(daiyunStem, "化權", palaces);
                string dyJi   = KbGetSiHuaPalace(daiyunStem, "化忌", palaces);
                if (!LfShouldSkipPalace(dyLu, age))
                {
                    if (keyPalaces.Contains(dyLu))   score += 10;
                    else if (goodPalaces.Contains(dyLu)) score += 6;
                    else if (!string.IsNullOrEmpty(dyLu)) score += 3;
                }
                if (!LfShouldSkipPalace(dyQuan, age))
                {
                    if (keyPalaces.Contains(dyQuan)) score += 6;
                    else if (!string.IsNullOrEmpty(dyQuan)) score += 3;
                }
                // 大限化忌不過濾（任何年齡都有意義）
                if (keyPalaces.Contains(dyJi))   score -= 18;
                else if (goodPalaces.Contains(dyJi)) score -= 10;
                else if (!string.IsNullOrEmpty(dyJi)) score -= 5;
            }
            return (int)Math.Round(Math.Clamp(score, 0, 100));
        }

        // 大限六親宮影響摘要：論對本命之助益或損失，考量年齡適切性，不談本命宮事項
        private static string DyDecadeLiuQinNote(string palaceLabel, int pStart, bool isEmpty)
        {
            string emptyNote = isEmpty ? "（空宮，需借對宮主星判斷大限助力強弱）" : "";
            string note = palaceLabel switch
            {
                "大限兄弟宮" =>
                    "此大限中手足平輩與同儕運勢：若宮星吉旺，大限可得朋輩助力、資源共享；" +
                    "若宮星不佳，則朋輩競爭或消耗較大，宜留意合夥與借貸關係。",
                "大限夫妻宮" when pStart < 35 =>
                    "此大限感情婚姻運：感情與婚姻為大限重點，宮星吉旺主緣份順利、伴侶得力；" +
                    "宮星不佳則感情多波折，宜多溝通、避免衝動決策。",
                "大限夫妻宮" when pStart < 55 =>
                    "此大限伴侶相處運：伴侶關係對大限整體影響顯著，宮星吉旺主相互支持、家庭穩固；" +
                    "宮星不佳則需多留意伴侶健康或感情摩擦，以穩定為要。",
                "大限夫妻宮" =>
                    "此大限晚年伴侶扶持運：宮星吉旺主伴侶陪伴得力、晚年生活品質佳；" +
                    "宮星不佳則需多關注伴侶健康，彼此照顧為要。",
                "大限子女宮" when pStart < 35 =>
                    "此大限子女緣份與晚輩下屬運：宮星吉旺主子嗣緣深，且晚輩、部屬、學生等均有助力；" +
                    "宮星不佳則子女緣薄，晚輩下屬亦難以依靠，宜多費心栽培。",
                "大限子女宮" =>
                    "此大限晚輩與下屬運：子女宮亦主晚輩、部屬、學生等後進關係；" +
                    "宮星吉旺主下屬得力、後進貢獻明顯，事業有人脈支撐；" +
                    "宮星不佳則部屬易生變故或難以信任，大限中需親力親為。",
                "大限僕役宮" =>
                    "此大限人際友誼與部屬運：宮星吉旺主社交廣博、得力助手多、貴人頻出；" +
                    "宮星不佳則需防小人、合夥易生糾紛，人際關係宜謹慎。",
                "大限父母宮" when pStart < 50 =>
                    "此大限長輩貴人與文書運：宮星吉旺主上司提攜、文件合約順利、印信得力；" +
                    "宮星不佳則長輩助力有限，凡事需自立，文書合約宜審慎。",
                "大限父母宮" =>
                    "此大限貴人文書與印信運：宮星吉旺主貴人相助、文件順利、名聲受保；" +
                    "宮星不佳則易有文書糾紛或貴人緣薄，大限中宜低調行事。",
                _ => $"此大限{palaceLabel.Replace("大限", "")}對整體大限運勢有間接影響，需結合宮干四化判斷助益或損失。"
            };
            return string.IsNullOrEmpty(emptyNote) ? note : note + emptyNote;
        }

        private static string DyCrossClass(int baziScore, int ziweiScore)
        {
            string bazi  = baziScore  >= 70 ? "喜" : baziScore  >= 50 ? "平" : "忌";
            string ziwei = ziweiScore >= 68 ? "吉" : ziweiScore >= 50 ? "平" : "凶";
            return (bazi, ziwei) switch
            {
                ("喜", "吉") => "大吉",
                ("喜", "平") or ("平", "吉") => "吉",
                ("喜", "凶") or ("平", "平") or ("忌", "吉") => "平",
                ("平", "凶") or ("忌", "平") => "小凶",
                ("忌", "凶") => "大凶",
                _ => "平"
            };
        }

        private static string DyYearSummary(string crossClass, string flStemSS, string flBranchSS, int baziScore, int ziweiScore)
        {
            string baziDesc  = baziScore  >= 70 ? "八字喜用" : baziScore  >= 50 ? "八字平和" : "八字忌神";
            string ziweiDesc = ziweiScore >= 68 ? "紫微吉臨" : ziweiScore >= 50 ? "紫微平穩" : "紫微凶曜";
            string bazi  = baziScore  >= 70 ? "喜" : baziScore  >= 50 ? "平" : "忌";
            string ziwei = ziweiScore >= 68 ? "吉" : ziweiScore >= 50 ? "平" : "凶";
            // 9種組合各給獨立建議，避免「平」一律守成的問題
            string action = (bazi, ziwei) switch
            {
                ("喜", "吉") => "雙重加持，此為難得機遇年，宜大膽佈局、積極進取。",
                ("喜", "平") => "八字得力為主，宜主動進取，穩中求進，趁勢而為。",
                ("平", "吉") => "紫微助力明顯，善用人際與外部資源，可借勢推進。",
                ("喜", "凶") => "八字雖好，紫微有阻，宜守好既有優勢，防外部因素干擾。",
                ("平", "平") => "八字紫微均屬平穩，無大起伏，靜守積累，蓄勢待發。",
                ("忌", "吉") => "八字承壓但紫微扶持，宜借重外力、人脈化解，少獨力硬幹。",
                ("平", "凶") => "紫微有阻，宜低調行事，避免不必要風險與正面衝突。",
                ("忌", "平") => "八字忌神為主要壓力，宜謹慎守成，減少重大決策。",
                ("忌", "凶") => "八字紫微雙重壓力，宜低調守成，防範風險，靜待轉機。",
                _            => "平穩行事，量力而為。"
            };
            // 極端分數加提示
            string extreme = (baziScore <= 22 && ziwei == "吉")
                ? "（八字壓力極重，紫微雖扶，仍需格外謹慎）"
                : (baziScore >= 80 && ziweiScore >= 75) ? "（得分極高，把握此難得好年）" : "";
            string ssDesc = !string.IsNullOrEmpty(flStemSS) ? $"（{flStemSS}年）" : "";
            return $"{baziDesc}、{ziweiDesc}，{crossClass}{ssDesc}。{action}{extreme}";
        }

        private static string DyCrossDesc(string crossClass, string flStemSS, string flBranchSS, int baziScore, int ziweiScore)
        {
            string bazi  = baziScore  >= 70 ? "喜" : baziScore  >= 50 ? "平" : "忌";
            string ziwei = ziweiScore >= 68 ? "吉" : ziweiScore >= 50 ? "平" : "凶";
            string desc = (bazi, ziwei) switch
            {
                ("喜", "吉") => "八字用神得力、紫微吉曜臨宮，雙重加持，為難得黃金年份。事業、投資、感情皆可大膽佈局，宜積極進取，乘勢而上，此種年份一生難得幾回。",
                ("喜", "平") => "八字喜用神活躍，八字面整體順暢；紫微宮位平穩，無特殊加分或減分。以八字優勢為主，宜主動進取，趁勢推進計畫，外部環境不阻力。",
                ("平", "吉") => "紫微吉曜加持明顯，外部機遇、人際貴人較多；八字本身平穩，無明顯阻力。宜把握外部帶來的機會，借勢推進，善用人脈與外部資源是關鍵。",
                ("喜", "凶") => "八字用神得力，自身狀態良好；惟紫微有凶曜干擾，外部環境或人際可能帶來摩擦。宜守好既有優勢，減少對外拓展，防範外部突發變數。",
                ("平", "平") => "八字與紫微均屬平穩中庸，無明顯吉凶加持。此年平中有機，宜靜守積累、鞏固基礎，不宜冒進，蓄勢待發，厚積薄發自有收穫。",
                ("忌", "吉") => $"八字忌神活躍（得分{baziScore}），自身承受一定壓力；幸有紫微吉曜扶持（得分{ziweiScore}），外部環境有助力。關鍵在於善用外部人脈資源化解內部壓力，切勿獨力硬幹，借勢借力可穩住局面。",
                ("平", "凶") => "紫微凶曜臨宮，外部干擾、意外變數較多；八字本身平穩，自身實力尚可。宜低調行事，減少對外衝突，專注內部鞏固，避免主動引發外部糾紛。",
                ("忌", "平") => $"八字忌神為主要壓力（得分{baziScore}），自身行事易有阻滯；紫微平穩，外部無特殊加減。宜謹慎守成，暫緩重大決策，以穩健低調行事為上策，減少犯錯機會。",
                ("忌", "凶") => $"八字忌神（得分{baziScore}）與紫微凶曜（得分{ziweiScore}）雙重壓力，內外均有阻礙。此年宜極度低調守成，切勿冒進或主動出擊，做好風險管理，靜待轉機，以守為攻。",
                _            => "平穩行事，量力而為，靜待機遇。"
            };
            string ssNote = "";
            if (!string.IsNullOrEmpty(flStemSS)) ssNote = $"流年天干{flStemSS}";
            if (!string.IsNullOrEmpty(flBranchSS)) ssNote += (ssNote.Length > 0 ? "、地支藏" : "地支藏") + flBranchSS;
            return string.IsNullOrEmpty(ssNote) ? desc : $"【{ssNote}】{desc}";
        }

        private static string DyBuildReport(
            string yStem, string yBranch, string mStem, string mBranch,
            string dStem, string dBranch, string hStem, string hBranch,
            string yStemSS, string mStemSS, string hStemSS,
            string yBranchSS, string mBranchSS, string dBranchSS, string hBranchSS,
            string dmElem, Dictionary<string, double> wuXing, double bodyPct, string bodyLabel,
            string season, string seaLabel, string pattern,
            string yongShenElem, string fuYiElem, string yongReason, string jiShenElem,
            List<(string stem, string branch, string liuShen, int startAge, int endAge)> luckCycles,
            List<(int year, int age, string flStem, string flBranch,
                string daiyunStem, string daiyunBranch,
                int baziScore, int ziweiScore, string crossClass,
                string flStemSS, string flBranchSS)> annualDetails,
            bool hasZiwei, JsonElement palaces,
            Dictionary<string, Dictionary<string, (string palace, string desc)>> siHuaDescMap,
            string ziweiFullContent, HashSet<string> chartStars,
            Dictionary<string, string> decadeKbMap,
            int gender, int birthYear, int years, string[] branches, string dStemRef)
        {
            var sb = new StringBuilder();
            string genderText = gender == 1 ? "男（乾造）" : "女（坤造）";
            string SS(string ss) => string.IsNullOrEmpty(ss) ? "" : $"（{ss}）";
            string wx = $"木{wuXing["木"]:F0}% 火{wuXing["火"]:F0}% 土{wuXing["土"]:F0}% 金{wuXing["金"]:F0}% 水{wuXing["水"]:F0}%";
            string yearsLabel = years == 0 ? "終身" : $"{years} 年";
            int startYear = annualDetails.Count > 0 ? annualDetails[0].year : DateTime.Today.Year;
            int endYear   = annualDetails.Count > 0 ? annualDetails[^1].year : startYear;

            sb.AppendLine("=================================================================");
            sb.AppendLine("                         大 運 命 書");
            sb.AppendLine("=================================================================");
            sb.AppendLine();

            // 人生指南目錄
            sb.AppendLine("                       人  生  指  南");
            sb.AppendLine("-----------------------------------------------------------------");
            sb.AppendLine("  命主資料與大運概況");
            sb.AppendLine("  格局與用神判定");
            sb.AppendLine("  分析期間大運干支論斷");
            sb.AppendLine("  流年逐年分析");
            sb.AppendLine("  重點宮位綜合評估");
            sb.AppendLine("  趨吉避凶總建議");
            sb.AppendLine("  人生警示事項（先天體質）");
            sb.AppendLine("-----------------------------------------------------------------");
            sb.AppendLine();

            // === Ch.1 命主資料 + 大運概況 ===
            sb.AppendLine("【第一章：命主資料與大運概況】");
            sb.AppendLine($"性別：{genderText}  出生年：{birthYear} 年");
            sb.AppendLine($"四柱：{yStem}{yBranch} {mStem}{mBranch} {dStem}{dBranch} {hStem}{hBranch}");
            sb.AppendLine($"十神：年干{SS(yStemSS)} 年支{SS(yBranchSS)} 月干{SS(mStemSS)} 月支{SS(mBranchSS)} 時干{SS(hStemSS)} 時支{SS(hBranchSS)}");
            sb.AppendLine($"日主：{dStem}（{dmElem}）  格局：{pattern}  日主{bodyLabel}（{bodyPct:F0}%）");
            sb.AppendLine($"用神：{yongShenElem}  忌神：{jiShenElem}  五行：{wx}");
            sb.AppendLine($"分析期間：{startYear} 至 {endYear} 年（{yearsLabel}大運命書）");
            int nowAge = DateTime.Today.Year - birthYear;
            var curLC = luckCycles.FirstOrDefault(lc => nowAge >= lc.startAge && nowAge < lc.endAge);
            if (!string.IsNullOrEmpty(curLC.stem))
            {
                string curLCSS  = LfStemShiShen(curLC.stem, dStemRef);
                string curLCBMs = LfBranchHiddenRatio.TryGetValue(curLC.branch, out var lcBh) && lcBh.Count > 0 ? lcBh[0].stem : "";
                string curLCBSS = !string.IsNullOrEmpty(curLCBMs) ? LfStemShiShen(curLCBMs, dStemRef) : "";
                sb.AppendLine($"當前大運：{curLC.stem}{curLC.branch}（天干{curLCSS}·地支{curLCBSS}），{curLC.startAge}-{curLC.endAge} 歲");
            }
            sb.AppendLine();

            // === Ch.2 格局與用神判定 ===
            sb.AppendLine("【第二章：格局與用神判定】");
            string tuneElemDisp2 = season == "冬" ? "火" : season == "夏" ? "水" : "";
            string jiYongElemDisp2 = LfElemOvercomeBy.GetValueOrDefault(yongShenElem, "");
            sb.AppendLine($"格局：【{pattern}】");
            sb.AppendLine($"用神：【{yongShenElem}】（{yongReason}）");
            sb.AppendLine($"喜用：天干 {LfElemStems(yongShenElem)}，地支 {LfElemBranches(yongShenElem)}");
            if (fuYiElem != yongShenElem)
                sb.AppendLine($"輔助喜神：【{fuYiElem}】（{(bodyPct <= 40 ? "印比互補扶身" : "官財互補制衡")}）");
            if (!string.IsNullOrEmpty(tuneElemDisp2) && tuneElemDisp2 != yongShenElem && tuneElemDisp2 != fuYiElem)
                sb.AppendLine($"調候喜神：【{tuneElemDisp2}】（{(season == "冬" ? "冬月寒凍，喜火暖局" : "夏月炎熱，喜水消暑")}）");
            sb.AppendLine($"大忌(X)：{jiShenElem}，天干 {LfElemStems(jiShenElem)}，地支 {LfElemBranches(jiShenElem)}");
            if (!string.IsNullOrEmpty(jiYongElemDisp2) && jiYongElemDisp2 != jiShenElem)
                sb.AppendLine($"次忌(△忌)：{jiYongElemDisp2}（克用神{yongShenElem}，力道較輕）");
            sb.AppendLine($"格局說明：{LfPatternDesc(pattern)}");
            sb.AppendLine();
            sb.AppendLine(LfBuildYongJiTable(yongShenElem, fuYiElem, jiShenElem, tuneElemDisp2, dStemRef, branches));
            sb.AppendLine();

            // === Ch.3 分析期間大運干支論斷 ===
            sb.AppendLine("【第三章：分析期間大運干支論斷】");
            string[] branchSSArr = { yBranchSS, mBranchSS, dBranchSS, hBranchSS };
            string[] dyChartStems = { yStem, mStem, dStemRef, hStem };
            string[] dyPillarBranches = { yBranch, mBranch, dBranch, hBranch };
            string[] dyDayEmpty = LfCalcDayEmpty(dStem, dBranch);
            string[] dyTiaoHouList = LfTiaoHou.TryGetValue(dStemRef, out var dyTh1) && dyTh1.TryGetValue(mBranch, out var dyTh2)
                ? dyTh2 : Array.Empty<string>();
            string dyTiaoHouElem = dyTiaoHouList.Length > 0 ? KbStemToElement(dyTiaoHouList[0]) : "";
            var coveredLucks = luckCycles.Where(lc =>
                annualDetails.Any(a => a.age >= lc.startAge && a.age < lc.endAge)).ToList();
            if (coveredLucks.Count == 0) coveredLucks = luckCycles.Take(2).ToList();
            // 分析期實際年齡範圍（限制大限宮位搜尋範圍，避免顯示分析期外的大限）
            int minAnalysisAge = annualDetails.Count > 0 ? annualDetails.Min(a => a.age) : int.MinValue;
            int maxAnalysisAge = annualDetails.Count > 0 ? annualDetails.Max(a => a.age) : int.MaxValue;
            foreach (var lc in coveredLucks)
            {
                string lcSS  = LfStemShiShen(lc.stem, dStemRef);
                string lcBMs = LfBranchHiddenRatio.TryGetValue(lc.branch, out var lcBH2) && lcBH2.Count > 0 ? lcBH2[0].stem : "";
                string lcBSS = !string.IsNullOrEmpty(lcBMs) ? LfStemShiShen(lcBMs, dStemRef) : "";
                int lcScore  = LfCalcLuckScore(lc.stem, lc.branch, pattern, yongShenElem, fuYiElem, jiShenElem,
                    dmElem, bodyPct > 50, dyTiaoHouElem, season, branches, dyChartStems, dStemRef);
                sb.AppendLine($"{lc.startAge}-{lc.endAge} 歲 大運：{lc.stem}{lc.branch}（天干{lcSS}·地支{lcBSS}）  評分：{lcScore} 分（{LfLuckLevel(lcScore)}）");
                sb.AppendLine($"  {LfLuckDesc(lcScore, LfLuckLevel(lcScore))}");
                // 天干事項：十神五行喜忌說明
                string lcStemElem = KbStemToElement(lc.stem);
                bool lcStemGood = lcStemElem == yongShenElem || lcStemElem == fuYiElem;
                bool lcStemBad  = lcStemElem == jiShenElem;
                string lcStemTrend = lcStemGood ? "屬喜用五行" : lcStemBad ? "屬忌神五行" : "屬中性五行";
                string lcStemEventDesc = lcSS switch
                {
                    "比肩" => lcStemGood ? "自立奮發，同輩互助，合夥共事有利。" : "競爭耗力，同輩牽制，宜各自獨立、防糾紛。",
                    "劫財" => lcStemGood ? "積極進取，破舊立新，有偏財機遇。" : "財務競爭激烈，宜防破財耗損、合夥是非，戒投機冒進。",
                    "食神" => lcStemGood ? "才藝展現，口福豐盛，子女緣佳，事業創作機會多。" : "耗洩過度，精力分散，需節制，防止才華難以變現。",
                    "傷官" => lcStemGood ? "才華外露，技術精進，適合創業突破舊局。" : "口舌是非多，易與上司對立，宜修身謙遜、防官司。",
                    "偏財" => lcStemGood ? "偏財運旺，父緣異性緣佳，廣結善緣有助財源。" : "財來財去，易衝動破財，宜謹慎理財、防詐騙。",
                    "正財" => lcStemGood ? "財運穩固，努力必有回報，婚姻穩定，適合穩健投資。" : "財庫受壓，勞而收穫有限，宜節流、保守理財。",
                    "七殺" => lcStemGood ? "壓力化為動力，可建功立業，適合競爭激烈的環境。" : "官非壓力大，健康情緒易受損，宜守成、防意外與衝突。",
                    "正官" => lcStemGood ? "名聲地位提升，升遷機會大，婚緣顯現。" : "規範束縛感強，職場壓力重，宜守紀律、防小人是非。",
                    "偏印" => lcStemGood ? "偏門學習進修，貴人助力，靈感豐富，適合研究鑽研。" : "思路偏執，食傷受制，宜廣納意見、防孤立封閉。",
                    "正印" => lcStemGood ? "印綬護身，學業晉升，長輩庇蔭，心靈沉穩有力。" : "依賴心重，行動力不足，宜主動出擊、防過度保守。",
                    _ => ""
                };
                if (!string.IsNullOrEmpty(lcStemEventDesc))
                    sb.AppendLine($"  【天干事項】大運天干{lc.stem}（{lcSS}），{lcStemTrend}：{lcStemEventDesc}");
                // 天干合：被命局天干合住說明
                if (LfTianGanHeMap.TryGetValue(lc.stem, out var lcTgHe) && dyChartStems.Contains(lcTgHe.stem))
                    sb.AppendLine($"  大運天干{lc.stem}被命局{lcTgHe.stem}合化，{(lcStemGood ? "喜用力量被牽制，效力略減。" : "忌神被合化，凶意稍緩。")}");
                // 地支事項
                string palaceEvents = LfBranchEventsPalace(lc.branch, lcBSS, branches, branchSSArr, lc.startAge);
                if (!string.IsNullOrEmpty(palaceEvents))
                {
                    sb.AppendLine($"  【地支事項】大運地支{lc.branch}（{lcBSS}）：");
                    sb.AppendLine($"  {palaceEvents}");
                }
                // 五大模組論斷（天干/地支引動/三干三支/空亡）
                string lcStepAnalysis = LfDyStepAnalysis(
                    lc.stem, lc.branch, lc.startAge, lc.endAge, nowAge,
                    dStemRef, dyChartStems, dyPillarBranches, branchSSArr,
                    yongShenElem, jiShenElem, dyDayEmpty, skipHeader: true);
                if (!string.IsNullOrEmpty(lcStepAnalysis)) sb.AppendLine(lcStepAnalysis);
                // 紫微大限宮干化忌入關鍵宮位警示
                if (hasZiwei)
                {
                    int lcAnalysisStart = Math.Max(lc.startAge, minAnalysisAge);
                    int lcAnalysisEnd   = Math.Min(lc.endAge,   maxAnalysisAge);
                    var warnPalaces = DyGetOverlappingDecadePalaces(palaces, lcAnalysisStart, lcAnalysisEnd);
                    var keyPalWarn  = new HashSet<string> { "命宮", "財帛宮", "官祿宮", "夫妻宮", "田宅宮" };
                    foreach (var (warnPalName, warnStem, warnPs, warnPe) in warnPalaces)
                    {
                        if (string.IsNullOrEmpty(warnStem)) continue;
                        if (!YearStemSiHuaMap.TryGetValue(warnStem, out var warnSiHua)) continue;
                        string dyJiPal = KbGetSiHuaPalace(warnStem, "化忌", palaces);
                        if (!string.IsNullOrEmpty(dyJiPal) && keyPalWarn.Contains(dyJiPal)
                            && !LfShouldSkipPalace(dyJiPal, lc.startAge))
                        {
                            string jiNote = dyJiPal switch
                            {
                                "命宮"   => "本命受衝，宜守護健康、防意外變故",
                                "財帛宮" => "財路受阻，整個大限需嚴防破財、詐騙、投資損失",
                                "田宅宮" => "財庫受損，整個大限需嚴防資產流失、詐騙、不動產糾紛",
                                "官祿宮" => "事業受阻，宜守成、防職場是非與官司",
                                "夫妻宮" => "婚姻感情有波折，宜多溝通、防感情變故",
                                _       => $"{dyJiPal}受影響，需謹慎應對"
                            };
                            sb.AppendLine($"  【紫微大限提醒】此段大限（{warnPs}-{warnPe}歲）宮干 {warnStem} 化忌入 {dyJiPal}：{jiNote}");
                        }
                    }
                }
                sb.AppendLine();
                // 大限宮位主星格局 + 四化（紫微大限宮位的宮干四化，非八字大運天干）
                // 八字大運可能橫跨多個紫微大限，逐段列出
                if (hasZiwei)
                {
                    int lcAnalysisStart2 = Math.Max(lc.startAge, minAnalysisAge);
                    int lcAnalysisEnd2   = Math.Min(lc.endAge,   maxAnalysisAge);
                    var overlapPalaces = DyGetOverlappingDecadePalaces(palaces, lcAnalysisStart2, lcAnalysisEnd2);
                    foreach (var (lcDecadePalace, lcDecadeStem, pStart, pEnd) in overlapPalaces)
                    {
                        int overlapStart = Math.Max(lc.startAge, pStart);
                        int overlapEnd   = Math.Min(lc.endAge, pEnd);
                        int overlapYears = overlapEnd - overlapStart + 1;
                        // 重疊不足2年（邊際跨限），只標記銜接提示，不展開完整大限分析
                        if (overlapYears < 2)
                        {
                            sb.AppendLine($"  ▍ 下一大限（{pStart}-{pEnd}歲）{lcDecadePalace}·宮干 {lcDecadeStem}（銜接提示，詳細分析見下一期大運命書）");
                            sb.AppendLine();
                            continue;
                        }
                        // 顯示宮位本身的完整大限範圍，括號內標注與八字大運重疊的歲數段
                        string decadeLabel = pStart == pEnd ? $"{pStart}歲" : $"{pStart}-{pEnd}歲";
                        string overlapNote = (overlapStart == pStart && overlapEnd == pEnd) ? ""
                            : $"（本大運涵蓋{overlapStart}-{overlapEnd}歲）";
                        sb.AppendLine($"  ▍ 大限（{decadeLabel}）{lcDecadePalace}·宮干 {lcDecadeStem}{overlapNote}");
                        sb.AppendLine($"  以{lcDecadePalace}為大限命宮，論{pStart}-{pEnd}歲大限格局：");
                        // 大限命宮格局：以該大限宮位地支為命宮查詢 KB，取「命宮」段落
                        string palaceStars = KbGetPalaceStars(palaces, lcDecadePalace);
                        string decadePalBranch = KbGetPalaceBranch(palaces, lcDecadePalace);
                        string decadeKbFull = decadeKbMap.GetValueOrDefault(decadePalBranch, "");
                        // 大限命宮格局：只傳宮位內星（不傳三方四正），避免三方四正星觸發會照描述
                        var lcDecPalStars = KbGetPalaceStarsSet(palaces, lcDecadePalace);
                        string palaceContent = KbFilterZiweiContent(
                            KbExtractPalaceSection(decadeKbFull, "命宮"),
                            lcDecPalStars,
                            lcDecPalStars);
                        // 移除尾端不完整的段落標題（過濾後星空，只剩「在三方四正中：」之類）
                        if (!string.IsNullOrEmpty(palaceContent))
                        {
                            var pcLines = palaceContent.Split('\n')
                                .Reverse()
                                .SkipWhile(l => l.TrimEnd().EndsWith("：") || string.IsNullOrWhiteSpace(l))
                                .Reverse()
                                .ToList();
                            palaceContent = string.Join("\n", pcLines).Trim();
                        }
                        if (!string.IsNullOrEmpty(palaceStars))
                        {
                            // 有主星：以大限命宮角度顯示主星 + 格局
                            sb.AppendLine($"  大限命宮主星：{palaceStars}");
                            // 生年化忌落入此大限宮位警示
                            if (YearStemSiHuaMap.TryGetValue(yStem, out var natalSiHuaDy))
                            {
                                string natalJiAbbr = natalSiHuaDy.ji;
                                string natalJiFull = StarAbbrToFull.TryGetValue(natalJiAbbr, out var jf) ? jf : natalJiAbbr;
                                var paStarsCheck = KbGetPalaceStarsSet(palaces, lcDecadePalace);
                                if (paStarsCheck.Contains(natalJiAbbr) || paStarsCheck.Contains(natalJiFull))
                                {
                                    // 大限宮位即為大限命宮，生年化忌坐大限命宮，格局先天受制
                                    string natalJiNote = "此星帶生年化忌坐大限命宮，整個大限格局先天受制，運勢易有起伏，行事宜謹慎務實，防範意外及突發變故";
                                    sb.AppendLine($"  【生年化忌】{natalJiFull}帶生年化忌（命主{yStem}年生），{natalJiNote}");
                                }
                            }
                            if (!string.IsNullOrEmpty(palaceContent))
                                sb.AppendLine($"  {palaceContent.Trim()}");
                        }
                        else
                        {
                            // 無主星：優先順序 六吉輔星 → 四剎星 → 格局名 → 空宮
                            var (goodAux, badAux) = DyGetPalaceAuxiliaryStars(palaces, lcDecadePalace);
                            if (goodAux.Count > 0)
                            {
                                sb.AppendLine($"  六吉輔星：{string.Join("、", goodAux)}");
                                foreach (var gs in goodAux)
                                {
                                    string note = DyAuxStarNote(gs);
                                    if (!string.IsNullOrEmpty(note))
                                        sb.AppendLine($"    {gs}：{note}");
                                }
                            }
                            if (badAux.Count > 0)
                            {
                                sb.AppendLine($"  四剎星：{string.Join("、", badAux)}");
                                foreach (var bs in badAux)
                                {
                                    string note = DyAuxStarNote(bs);
                                    if (!string.IsNullOrEmpty(note))
                                        sb.AppendLine($"    {bs}：{note}");
                                }
                            }
                            // 格局名或非星條件的通用描述（KbFilterZiweiContent 通用格局已通過篩選）
                            if (!string.IsNullOrEmpty(palaceContent))
                                sb.AppendLine($"  {palaceContent.Trim()}");
                            // 大限財宮 / 大限事業宮（+4/+8 地支偏移）
                            var (finPalace, finStars, carPalace, carStars) = DyGetDecadeKeyPalaceStars(palaces, lcDecadePalace);
                            if (!string.IsNullOrEmpty(finPalace))
                                sb.AppendLine($"  大限財宮：{finPalace}（{(string.IsNullOrEmpty(finStars) ? "空宮" : finStars)}）");
                            if (!string.IsNullOrEmpty(carPalace) && !LfShouldSkipPalace("官祿宮", pStart))
                                sb.AppendLine($"  大限事業宮：{carPalace}（{(string.IsNullOrEmpty(carStars) ? "空宮" : carStars)}）");
                            if (goodAux.Count == 0 && badAux.Count == 0 && string.IsNullOrEmpty(palaceContent))
                                sb.AppendLine($"  此宮為空宮，需借對宮主星論斷，吉凶受三方四正飛化影響為主。");
                        }
                        sb.AppendLine();
                        // 宮干四化
                        if (!string.IsNullOrEmpty(lcDecadeStem) && siHuaDescMap.TryGetValue(lcDecadeStem, out var lcDyMap)
                            && YearStemSiHuaMap.TryGetValue(lcDecadeStem, out var lcSiHua))
                        {
                            sb.AppendLine($"  宮干 {lcDecadeStem} 四化：");
                            string[] dyShTypes = { "化祿", "化權", "化科", "化忌" };
                            string[] dyStars   = { lcSiHua.lu, lcSiHua.quan, lcSiHua.ke, lcSiHua.ji };
                            for (int si = 0; si < dyShTypes.Length; si++)
                            {
                                var (pal, desc) = lcDyMap.GetValueOrDefault(dyShTypes[si], ("", ""));
                                if (!string.IsNullOrEmpty(pal) && LfShouldSkipPalace(pal, pStart)) continue;
                                string palLabel = string.IsNullOrEmpty(pal) ? "（命盤未含此星）" : $"入{pal}";
                                string descText = string.IsNullOrEmpty(desc) ? "" : $"　{desc}";
                                sb.AppendLine($"  {dyShTypes[si]}（{dyStars[si]}星）{palLabel}{descText}");
                            }
                        }
                        sb.AppendLine();
                        // 大限十二宮星化象（以大限宮位為大限命宮，逆時針展開其餘十一宮）
                        if (!string.IsNullOrEmpty(decadeKbFull))
                        {
                            sb.AppendLine($"  【大限十二宮星化象（以{lcDecadePalace}為大限命宮）】");
                            string[] dyPalLabels = {
                                "大限命宮", "大限兄弟宮", "大限夫妻宮", "大限子女宮",
                                "大限財帛宮", "大限疾厄宮", "大限遷移宮", "大限僕役宮",
                                "大限官祿宮", "大限田宅宮", "大限福德宮", "大限父母宮" };
                            string[] dbSections = {
                                "命宮", "兄弟宮", "夫妻宮", "子女宮",
                                "財帛宮", "疾厄宮", "遷移宮", "交友宮",
                                "事業宮", "田宅宮", "福德宮", "父母宮" };
                            // 六親宮 offsets：兄弟(1)/夫妻(2)/子女(3)/僕役(7)/父母(11)
                            // 六親宮不使用 KB 文字（均為本命宮論法，談幾子幾女等），改以大限影響摘要
                            var liuQinOffsets = new HashSet<int> { 1, 2, 3, 7, 11 };
                            for (int po = 1; po < 12; po++) // offset 0 = 大限命宮，已在上方顯示
                            {
                                string ccwBranch    = DyGetCCWBranch(decadePalBranch, po);
                                string natalPalName = KbGetPalaceNameByBranch(palaces, ccwBranch);
                                string natalStars   = KbGetPalaceStars(palaces, natalPalName);
                                string starsLabel   = string.IsNullOrEmpty(natalStars) ? "空宮" : natalStars;
                                string natalLabel   = string.IsNullOrEmpty(natalPalName) ? "" : $"（本命{natalPalName}）";
                                sb.AppendLine($"  {dyPalLabels[po]}{natalLabel}：{starsLabel}");
                                if (liuQinOffsets.Contains(po))
                                {
                                    // 六親宮：論大限對六親關係的助益/損失，不談本命事項，需考量年齡適切性
                                    string liuQinNote = DyDecadeLiuQinNote(dyPalLabels[po], pStart, string.IsNullOrEmpty(natalStars));
                                    sb.AppendLine($"  {liuQinNote}");
                                }
                                else
                                {
                                    // 非六親宮（財/疾/遷/官/田/福）：顯示 KB 星性描述
                                    var natalStarsSet  = KbGetPalaceStarsSet(palaces, natalPalName);
                                    string sectionRaw  = KbExtractPalaceSection(decadeKbFull, dbSections[po]);
                                    string sectionFilt = KbFilterZiweiContent(sectionRaw, natalStarsSet, natalStarsSet);
                                    if (!string.IsNullOrEmpty(sectionFilt))
                                    {
                                        var pcLines = sectionFilt.Split('\n')
                                            .Reverse()
                                            .SkipWhile(l => l.TrimEnd().EndsWith("：") || string.IsNullOrWhiteSpace(l))
                                            .Reverse()
                                            .ToList();
                                        sectionFilt = string.Join("\n", pcLines).Trim();
                                        sb.AppendLine($"  {sectionFilt.Replace("\n", "\n  ").Trim()}");
                                    }
                                }
                                sb.AppendLine();
                            }
                        }
                    }
                }
                sb.AppendLine();
            }
            sb.AppendLine();

            // === Ch.4 流年逐年分析 ===
            sb.AppendLine("【第四章：流年逐年分析】");
            if (!hasZiwei)
                sb.AppendLine("（命盤無紫微資料，紫微評分以預設值呈現，八字分析仍完整。）");
            // 年齡適切性提示（依分析期間起始年齡）
            {
                int ch4Age = annualDetails.Count > 0 ? annualDetails[0].age : 0;
                string ageHint = LfAgeTopicHint(ch4Age);
                if (!string.IsNullOrEmpty(ageHint)) sb.AppendLine(ageHint);
            }
            sb.AppendLine();
            foreach (var d in annualDetails)
            {
                sb.AppendLine($"【{d.year} 年】流年：{d.flStem}{d.flBranch}  年齡：{d.age} 歲  大運：{d.daiyunStem}{d.daiyunBranch}");
                sb.AppendLine($"  八字評分：{d.baziScore} 分  紫微評分：{d.ziweiScore} 分  綜合：【{d.crossClass}】");
                sb.AppendLine();

                // 八字面向
                sb.AppendLine("  ▍ 八字面向");
                string flStemElem = KbStemToElement(d.flStem);
                string flBrMs = LfBranchHiddenRatio.TryGetValue(d.flBranch, out var flBrh) && flBrh.Count > 0 ? flBrh[0].stem : "";
                string flBrElem = !string.IsNullOrEmpty(flBrMs) ? KbStemToElement(flBrMs) : "";
                bool stemGood = flStemElem == yongShenElem || flStemElem == fuYiElem;
                bool stemBad  = flStemElem == jiShenElem;
                bool brGood   = flBrElem == yongShenElem || flBrElem == fuYiElem;
                bool brBad    = flBrElem == jiShenElem;
                string stemCls = stemBad ? "X（大忌）" : stemGood ? "○（喜用）" : "△（中性）";
                string brCls   = string.IsNullOrEmpty(flBrElem) ? "△（中性）"
                    : brBad ? "X（大忌）" : brGood ? "○（喜用）" : "△（中性）";
                // 十神白話說明
                string DyStemNote(string ss, bool good, bool bad) => ss switch
                {
                    "比肩" => good ? "自立奮發，同輩合作有利" : bad ? "競爭耗力，防同輩或兄弟財務糾紛" : "有合作機會，需防競爭",
                    "劫財" => good ? "積極進取，有偏財機遇" : bad ? "破財競爭，戒衝動投資，防財務損耗" : "有活動能量，謹慎理財",
                    "食神" => good ? "才藝展現，口福豐盛，子女緣佳" : bad ? "耗洩過度，精力分散，防才華難變現" : "創作機會多，留意體力",
                    "傷官" => good ? "才華外露，技術精進，適合突破創新" : bad ? "口舌是非多，易與上司衝突，防官司" : "有創意但需謹言慎行",
                    "偏財" => good ? "偏財運旺，廣結善緣，父緣異性緣佳" : bad ? "財來財去，衝動破財，防詐騙" : "有財機，宜保守理財",
                    "正財" => good ? "財運穩固，努力有回報，婚姻穩定" : bad ? "財庫受壓，勞而少收，宜節流" : "財務平穩，踏實耕耘",
                    "七殺" => good ? "壓力化動力，可建功立業" : bad ? "官非壓力大，防意外衝突，宜守成" : "有競爭壓力，保持冷靜",
                    "正官" => good ? "名聲地位提升，升遷機會，婚緣顯現" : bad ? "規範束縛感強，防小人是非" : "職場有規範要求，守分則吉",
                    "偏印" => good ? "偏門學習，貴人助力，靈感豐富" : bad ? "思路偏執，防孤立，廣納意見" : "適合研究進修",
                    "正印" => good ? "長輩庇蔭，學業晉升，心靈沉穩" : bad ? "依賴心重，行動力不足，主動出擊" : "宜進修學習，穩健前行",
                    _ => ""
                };
                string stemNote = DyStemNote(d.flStemSS, stemGood, stemBad);
                string brNote   = !string.IsNullOrEmpty(d.flBranchSS) ? DyStemNote(d.flBranchSS, brGood, brBad) : "";
                sb.AppendLine($"  流年天干 {d.flStem}（{d.flStemSS}·{stemCls}）{(string.IsNullOrEmpty(stemNote) ? "" : "：" + stemNote)}");
                if (!string.IsNullOrEmpty(flBrElem))
                    sb.AppendLine($"  流年地支 {d.flBranch}（{d.flBranchSS}·{brCls}）{(string.IsNullOrEmpty(brNote) ? "" : "：" + brNote)}");
                // 歲君互動白話翻譯
                string brEvents = LfBranchEvents(d.flBranch, branches);
                if (!string.IsNullOrEmpty(brEvents))
                {
                    // 把技術術語轉成白話
                    string eventsNote = brEvents
                        .Replace("六沖", "對沖命局，變動動盪較大")
                        .Replace("六合", "與命局合化，")
                        .Replace("三會", "三會局，五行集中，")
                        .Replace("三合", "三合局，")
                        .Replace("三刑", "三刑沖局，是非爭訟健康需注意")
                        .Replace("六害", "暗中阻礙，人際摩擦需留意")
                        .Replace("六破", "事情有始無終，防財物耗損");
                    sb.AppendLine($"  歲君互動：{eventsNote}");
                }
                sb.AppendLine();

                // 紫微面向
                sb.AppendLine("  ▍ 紫微面向");
                if (YearStemSiHuaMap.TryGetValue(d.flStem, out var siHua))
                {
                    string[] flShTypes = { "化祿", "化權", "化科", "化忌" };
                    string[] flStars   = { siHua.lu, siHua.quan, siHua.ke, siHua.ji };
                    string decadeMingBr = hasZiwei ? DyGetDecadeMingBranch(palaces, d.age) : "";
                    if (hasZiwei && siHuaDescMap.TryGetValue(d.flStem, out var flDyMap))
                    {
                        for (int si = 0; si < flShTypes.Length; si++)
                        {
                            var (pal, desc) = flDyMap.GetValueOrDefault(flShTypes[si], ("", ""));
                            if (!string.IsNullOrEmpty(pal) && LfShouldSkipPalace(pal, d.age)) continue;
                            string decadePal = "";
                            if (!string.IsNullOrEmpty(pal) && !string.IsNullOrEmpty(decadeMingBr))
                            {
                                string starBr = KbGetBranchByPalaceName(palaces, pal);
                                decadePal = DyGetDecadePalaceName(starBr, decadeMingBr);
                            }
                            string palLabel = string.IsNullOrEmpty(pal) ? "（命盤未含此星）"
                                : string.IsNullOrEmpty(decadePal) ? $"入本命{pal}"
                                : $"入本命{pal}（大限{decadePal}）";
                            string descText = string.IsNullOrEmpty(desc) ? "" : $"：{desc}";
                            string decadeNote = string.IsNullOrEmpty(decadePal) ? ""
                                : DyGetDecadeSiHuaNote(flStars[si], flShTypes[si], decadePal);
                            string decadeNoteText = string.IsNullOrEmpty(decadeNote) ? "" : $"\n    ↳ 大限{decadePal}：{decadeNote}。";
                            sb.AppendLine($"  {flShTypes[si]}（{flStars[si]}星）{palLabel}{descText}{decadeNoteText}");
                        }
                    }
                    else if (hasZiwei)
                    {
                        for (int si = 0; si < flShTypes.Length; si++)
                        {
                            string pal = KbGetSiHuaPalace(d.flStem, flShTypes[si], palaces);
                            if (!string.IsNullOrEmpty(pal) && LfShouldSkipPalace(pal, d.age)) continue;
                            string decadePal = "";
                            if (!string.IsNullOrEmpty(pal) && !string.IsNullOrEmpty(decadeMingBr))
                            {
                                string starBr = KbGetBranchByPalaceName(palaces, pal);
                                decadePal = DyGetDecadePalaceName(starBr, decadeMingBr);
                            }
                            string palLabel = string.IsNullOrEmpty(pal) ? "（命盤未含此星）"
                                : string.IsNullOrEmpty(decadePal) ? $"入本命{pal}"
                                : $"入本命{pal}（大限{decadePal}）";
                            string decadeNote = string.IsNullOrEmpty(decadePal) ? ""
                                : DyGetDecadeSiHuaNote(flStars[si], flShTypes[si], decadePal);
                            string decadeNoteText = string.IsNullOrEmpty(decadeNote) ? "" : $"\n    ↳ 大限{decadePal}：{decadeNote}。";
                            sb.AppendLine($"  {flShTypes[si]}（{flStars[si]}星）{palLabel}{decadeNoteText}");
                        }
                    }
                    else
                        sb.AppendLine($"  流年四化：化祿（{siHua.lu}）、化權（{siHua.quan}）、化科（{siHua.ke}）、化忌（{siHua.ji}）");
                }
                sb.AppendLine();

                // 綜合論斷
                sb.AppendLine("  ▍ 綜合論斷");
                sb.AppendLine($"  {DyCrossDesc(d.crossClass, d.flStemSS, d.flBranchSS, d.baziScore, d.ziweiScore)}");
                sb.AppendLine();
                // 五大模組論斷（流年版）
                string lnStepAnalysis = LfLnYearAnalysis(
                    d.flStem, d.flBranch, d.year, LfBranchZodiac.GetValueOrDefault(d.flBranch, ""),
                    dStemRef, dyChartStems, dyPillarBranches, branchSSArr,
                    yongShenElem, jiShenElem, dyDayEmpty,
                    dayunBranch: d.daiyunBranch, skipHeader: true);
                if (!string.IsNullOrEmpty(lnStepAnalysis)) sb.AppendLine(lnStepAnalysis);
                sb.AppendLine("  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -");
                sb.AppendLine();
            }

            // === Ch.5 重點宮位綜合 ===
            sb.AppendLine("【第五章：重點宮位綜合評估】");
            if (annualDetails.Count > 0)
            {
                double avgBazi  = annualDetails.Average(a => (double)a.baziScore);
                double avgZiwei = annualDetails.Average(a => (double)a.ziweiScore);
                var bestYears = annualDetails.Where(a => a.crossClass is "大吉" or "吉").OrderByDescending(a => a.baziScore + a.ziweiScore).Take(3).ToList();
                var badYears  = annualDetails.Where(a => a.crossClass is "大凶" or "小凶").OrderBy(a => a.baziScore + a.ziweiScore).Take(3).ToList();
                sb.AppendLine($"分析期間八字平均：{avgBazi:F0} 分  紫微平均：{avgZiwei:F0} 分");
                if (bestYears.Count > 0)
                {
                    sb.AppendLine("重點吉年（宜把握）：");
                    foreach (var y in bestYears)
                        sb.AppendLine($"  {y.year} 年（{y.flStem}{y.flBranch}，{y.crossClass}，八字{y.baziScore}分·紫微{y.ziweiScore}分）");
                }
                if (badYears.Count > 0)
                {
                    sb.AppendLine("重點凶年（宜謹慎）：");
                    foreach (var y in badYears)
                        sb.AppendLine($"  {y.year} 年（{y.flStem}{y.flBranch}，{y.crossClass}，八字{y.baziScore}分·紫微{y.ziweiScore}分）");
                }
                sb.AppendLine();
                sb.AppendLine($"財帛宮方向：用神 {yongShenElem} 旺年財運較佳，忌神 {jiShenElem} 旺年需謹守財務，避免大額投資。");
                sb.AppendLine($"事業官祿：格局 {pattern}，天生適合{LfCareerDesc(pattern)}，喜用年宜積極進取，凶年宜守成。");
                string spouseStar = gender == 1 ? "妻星（財）" : "夫星（官殺）";
                sb.AppendLine($"夫妻感情：{spouseStar} 活躍年份感情變動較大，日支 {dBranch} 逢沖年份宜多溝通。");
                sb.AppendLine($"健康疾厄：{LfHealthDesc(wuXing, seaLabel).Split('\n')[0].Trim()}");
            }
            sb.AppendLine();

            // === Ch.6 趨吉避凶 ===
            sb.AppendLine("【第六章：趨吉避凶總建議】");
            if (annualDetails.Count > 0)
            {
                var allGood = annualDetails.Where(a => a.crossClass is "大吉" or "吉").ToList();
                var allBad  = annualDetails.Where(a => a.crossClass is "大凶" or "小凶").ToList();
                if (allGood.Count > 0)
                {
                    sb.AppendLine("最佳把握年份：");
                    sb.Append("  ");
                    sb.AppendLine(string.Join("、", allGood.Select(y => $"{y.year} 年（{y.crossClass}）")));
                    sb.AppendLine("  此類年份宜積極展開事業佈局、感情推進、投資理財。");
                }
                if (allBad.Count > 0)
                {
                    sb.AppendLine("需謹慎年份：");
                    sb.Append("  ");
                    sb.AppendLine(string.Join("、", allBad.Select(y => $"{y.year} 年（{y.crossClass}）")));
                    sb.AppendLine("  此類年份宜低調保守，避免重大決策，做好財務風險管理，注意健康。");
                }
            }
            sb.AppendLine();
            sb.AppendLine("具體行動建議：");
            sb.AppendLine($"  1. 喜用方位：{LfElemDir(yongShenElem)}，喜用色彩：{LfElemColor(yongShenElem)}");
            sb.AppendLine($"  2. 忌諱方向：{LfElemDir(jiShenElem)}，避免過多接觸 {jiShenElem} 屬性的人事物");
            sb.AppendLine($"  3. 事業宜從事：{LfElemCareer(yongShenElem)}");
            sb.AppendLine($"  4. 八字與紫微雙吉年宜大膽把握，雙凶年宜蟄伏蓄積，一吉一凶年宜穩健行事");
            sb.AppendLine();
            // === Ch.7 人生警示事項 ===
            sb.AppendLine("【第七章：人生警示事項（先天體質）】");
            sb.AppendLine();
            sb.AppendLine("▍ 小人防範");
            sb.AppendLine(LfXiaoRenAnalysis(yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch, jiShenElem, dmElem));
            sb.AppendLine();
            sb.AppendLine("▍ 官司文書風險");
            sb.AppendLine(LfGuanSiAnalysis(yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch, jiShenElem, dmElem, bodyPct));
            sb.AppendLine();
            sb.AppendLine("▍ 車關時機");
            {
                string cheGuanBase = LfCheGuanAnalysis(yBranch, mBranch, dBranch, hBranch, jiShenElem, dmElem);
                sb.AppendLine(cheGuanBase);
                // 掃描分析期內的引動年份
                string guanShaElemCg = LfElemOvercomeBy.GetValueOrDefault(dmElem, "");
                var siYiMaCg = new HashSet<string> { "寅", "申", "巳", "亥" };
                var riskBranchesCg = new[] { yBranch, mBranch, dBranch, hBranch }
                    .Where(b => siYiMaCg.Contains(b) && LfBranchElem.GetValueOrDefault(b, "") == guanShaElemCg)
                    .ToList();
                if (riskBranchesCg.Count > 0)
                {
                    var cheGuanYears = annualDetails
                        .Where(a => riskBranchesCg.Any(rb =>
                            LfChong.Contains(rb + a.flBranch) || LfChong.Contains(a.flBranch + rb)))
                        .Select(a => a.year).ToList();
                    if (cheGuanYears.Count > 0)
                        sb.AppendLine($"分析期間車關警示年份：{string.Join("、", cheGuanYears.Select(y => $"{y}年"))}（流年沖動驛馬官殺）");
                }
            }
            sb.AppendLine();
            sb.AppendLine("▍ 海外發展");
            sb.AppendLine(LfHaiWaiAnalysis(yBranch, mBranch, dBranch, hBranch, yongShenElem, jiShenElem, dmElem, hasZiwei, palaces));
            sb.AppendLine();
            sb.AppendLine("▍ 天乙貴人方向");
            {
                var tianYiMapDy = new Dictionary<string, string>
                {
                    {"甲","丑未"},{"戊","丑未"},{"庚","丑未"},
                    {"乙","子申"},{"己","子申"},
                    {"丙","亥酉"},{"丁","亥酉"},
                    {"壬","卯巳"},{"癸","卯巳"},
                    {"辛","午寅"}
                };
                string tianYiBranchesDy = tianYiMapDy.GetValueOrDefault(dStem, "");
                sb.AppendLine($"{dStem} 日主，天乙貴人在：{tianYiBranchesDy}（見此地支方位或行此地支大運年份，貴人助力最強）");
            }
            sb.AppendLine();

            sb.AppendLine("-----------------------------------------------------------------");
            sb.AppendLine("命理大師：玉洞子 | 大運命書 v1.2");
            return sb.ToString();
        }

        // === Ln (流年命書) Endpoint ===

        [HttpGet("analyze-liunian")]
        [Authorize]
        public async Task<IActionResult> GetLiunianAnalysis([FromQuery] int year = 0)
        {
            var identity = User.FindFirstValue(ClaimTypes.Email)
                         ?? User.FindFirstValue(ClaimTypes.Name)
                         ?? User.FindFirst("unique_name")?.Value;
            if (string.IsNullOrEmpty(identity))
                return Unauthorized(new { error = "請重新登入" });

            var user = await _context.Users.FirstOrDefaultAsync(u => u.UserName == identity || u.Email == identity);
            if (user == null) return BadRequest(new { error = "找不到用戶" });

            int currentYear = DateTime.Today.Year;
            if (year == 0) year = currentYear;
            if (year < currentYear)
                return BadRequest(new { error = $"流年命書不支援過去年份，請選擇 {currentYear} 年或以後" });

            var userChart = await _context.UserCharts.FirstOrDefaultAsync(c => c.UserId == user.Id);
            if (userChart == null || string.IsNullOrEmpty(userChart.ChartJson))
                return BadRequest(new { error = "no_chart" });

            bool lnIsAdmin = string.Equals(user.Email, _config["Admin:Email"], StringComparison.OrdinalIgnoreCase);
            int lnSubId = -1;
            if (!lnIsAdmin)
            {
                var (lnOk, lnErr, lnSubIdVal) = await CheckSubscriptionQuota(user.Id, "BOOK_LIUNIAN");
                if (!lnOk) return BadRequest(new { error = lnErr });
                lnSubId = lnSubIdVal;
            }

            try
            {
                var root = JsonDocument.Parse(userChart.ChartJson).RootElement;
                if (!root.TryGetProperty("bazi", out var bazi) && !root.TryGetProperty("baziInfo", out bazi))
                    return BadRequest(new { error = "命盤資料格式錯誤" });

                var yearP  = LfGetPillar(bazi, "yearPillar");
                var monthP = LfGetPillar(bazi, "monthPillar");
                var dayP   = LfGetPillar(bazi, "dayPillar");
                var timeP  = LfGetPillar(bazi, "timePillar");

                string yStem = LfPillarStem(yearP);   string yBranch = LfPillarBranch(yearP);
                string mStem = LfPillarStem(monthP);  string mBranch = LfPillarBranch(monthP);
                string dStem = LfPillarStem(dayP);    string dBranch = LfPillarBranch(dayP);
                string hStem = LfPillarStem(timeP);   string hBranch = LfPillarBranch(timeP);

                string yStemSS   = LfPillarStemSS(yearP);
                string mStemSS   = LfPillarStemSS(monthP);
                string hStemSS   = LfPillarStemSS(timeP);
                string yBranchSS = LfPillarBranchMainSS(yearP);
                string mBranchSS = LfPillarBranchMainSS(monthP);
                string dBranchSS = LfPillarBranchMainSS(dayP);
                string hBranchSS = LfPillarBranchMainSS(timeP);

                int birthYear = user.BirthYear ?? (DateTime.Today.Year - 30);
                int gender    = user.BirthGender ?? 1;
                string dmElem = KbStemToElement(dStem);
                var branches  = new[] { yBranch, mBranch, dBranch, hBranch };
                var wuXing    = LfCalcWuXingMatrix(yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch);
                double bodyPct   = LfGetBodyStrengthPct(dmElem, wuXing);
                string bodyLabel = LfGetBodyStrengthLabel(bodyPct);
                string season    = LfGetSeason(mBranch);
                string seaLabel  = LfGetSeasonLabel(mBranch);
                var (pattern, yongShenElem, fuYiElem, yongReason, tiaoHouElem) = LfDetectGeJuAndYongShen(
                    yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch,
                    dmElem, wuXing, bodyPct, season);
                string jiShenElem = LfGetJiShenElem(yongShenElem, dmElem, bodyPct, pattern);
                var chartStems4   = new[] { yStem, mStem, dStem, hStem };

                var luckCycles = LfExtractLuckCycles(root);
                bool hasZiwei  = root.TryGetProperty("palaces", out var palaces) && palaces.ValueKind == JsonValueKind.Array;

                var (flStem, flBranch) = DyGetYearStemBranch(year);
                int flAge = year - birthYear;
                var curLuck = luckCycles.FirstOrDefault(lc => flAge >= lc.startAge && flAge < lc.endAge);
                string daiyunStem   = curLuck.stem   ?? "";
                string daiyunBranch = curLuck.branch ?? "";
                string daiyunSS     = !string.IsNullOrEmpty(daiyunStem) ? LfStemShiShen(daiyunStem, dStem) : "";
                string daiyunBrMs   = LfBranchHiddenRatio.TryGetValue(daiyunBranch, out var dyBrH) && dyBrH.Count > 0 ? dyBrH[0].stem : "";
                string daiyunBrSS   = !string.IsNullOrEmpty(daiyunBrMs) ? LfStemShiShen(daiyunBrMs, dStem) : "";

                int flBaziScore  = DyCalcFlowYearBaziScore(flStem, flBranch, pattern, yongShenElem, fuYiElem, jiShenElem,
                    dmElem, bodyPct > 50, tiaoHouElem, season, branches, chartStems4);
                int flZiweiScore = hasZiwei ? DyCalcZiweiScore(flStem, palaces, daiyunStem, flAge) : 50;
                string flCrossClass = DyCrossClass(flBaziScore, flZiweiScore);
                string flStemSS   = LfStemShiShen(flStem, dStem);
                string flBrMs     = LfBranchHiddenRatio.TryGetValue(flBranch, out var flBrH) && flBrH.Count > 0 ? flBrH[0].stem : "";
                string flBranchSS = !string.IsNullOrEmpty(flBrMs) ? LfStemShiShen(flBrMs, dStem) : "";

                string shengXiaoBranch = DyGetYearStemBranch(birthYear).branch;

                // 預取四化（流年干 + 各月干）
                var siHuaDescMap = new Dictionary<string, Dictionary<string, (string palace, string desc)>>();
                if (hasZiwei)
                {
                    var uniqueStems = new List<string> { flStem };
                    for (int m = 1; m <= 12; m++)
                    {
                        var (ms, _) = LnGetMonthStemBranch(flStem, m);
                        if (!uniqueStems.Contains(ms)) uniqueStems.Add(ms);
                    }
                    string[] siHuaTypes = { "化祿", "化權", "化科", "化忌" };
                    foreach (var stem in uniqueStems.Where(s => YearStemSiHuaMap.ContainsKey(s)))
                    {
                        var stemMap = new Dictionary<string, (string palace, string desc)>();
                        foreach (var sh in siHuaTypes)
                        {
                            string pal  = KbGetSiHuaPalace(stem, sh, palaces);
                            string desc = string.IsNullOrEmpty(pal) ? "" : await KbSiHuaQuery(stem, sh, palaces);
                            if (desc.Length > 60) { int dot = desc.IndexOfAny(new[] { '。', '，', '\n' }); desc = dot > 0 && dot < 80 ? desc[..(dot + 1)] : desc[..60] + "..."; }
                            stemMap[sh] = (pal, desc);
                        }
                        siHuaDescMap[stem] = stemMap;
                    }
                }

                // 逐月分析
                string[] monthBranchLabels = { "寅月(約2月)","卯月(約3月)","辰月(約4月)","巳月(約5月)","午月(約6月)","未月(約7月)","申月(約8月)","酉月(約9月)","戌月(約10月)","亥月(約11月)","子月(約12月)","丑月(約1月)" };
                var monthlyDetails   = new List<(int idx, string mStemM, string mBranchM, string mSeason, int bazi, int ziwei, string cross, string flowStar, string tip)>();
                var monthlyForecasts = new List<object>();
                for (int m = 1; m <= 12; m++)
                {
                    var (mStemM, mBranchM) = LnGetMonthStemBranch(flStem, m);
                    string mSeason    = LnGetMonthSeason(m);
                    int    mBazi      = DyCalcFlowYearBaziScore(mStemM, mBranchM, pattern, yongShenElem, fuYiElem, jiShenElem,
                        dmElem, bodyPct > 50, tiaoHouElem, mSeason, branches, chartStems4);
                    int    mZiwei     = hasZiwei ? DyCalcZiweiScore(mStemM, palaces, daiyunStem, flAge) : 50;
                    string mCross     = DyCrossClass(mBazi, mZiwei);
                    string flowStar   = hasZiwei ? LnGetMonthFlowStar(m, flBranch, palaces) : "";
                    string mStemSSM   = LfStemShiShen(mStemM, dStem);
                    string mStemElemM = KbStemToElement(mStemM);
                    string baziHint   = (mStemElemM == yongShenElem || mStemElemM == fuYiElem) ? $"月干{mStemM}({mStemSSM})喜用"
                                      : mStemElemM == jiShenElem ? $"月干{mStemM}({mStemSSM})忌神" : $"月干{mStemM}({mStemSSM})中性";
                    string tip        = LnMonthTip(mCross, mSeason, mStemSSM);
                    monthlyDetails.Add((m, mStemM, mBranchM, mSeason, mBazi, mZiwei, mCross, flowStar, tip));
                    monthlyForecasts.Add(new {
                        month = m, label = monthBranchLabels[m - 1], stemBranch = mStemM + mBranchM,
                        season = mSeason, flowStar, baziHint, crossClass = mCross,
                        baziScore = mBazi, ziweiScore = mZiwei, tip
                    });
                }

                // 五術計算：命宮地支（純八字公式：月柱地支 + 時柱地支）
                string fallbackMingBranch = LnCalcMingBranchFromBazi(mBranch, hBranch);

                string taisuiGuard = LnCalcMingGongGuard(flBranch, hasZiwei ? palaces : default, fallbackMingBranch);
                var (taisuiRelation, taisuiLevel, taisuiDesc, taisuiNeedAn, taisuiPos) = LnCalcTaisuiRelation(shengXiaoBranch, flBranch);
                string taisuiGen = SixtyYearTaisuiGen.GetValueOrDefault(flStem + flBranch, "");
                string taisuiBranchEvent = BranchTaisuiEvent.GetValueOrDefault(flBranch, "");
                string blindSect = LnCalcBlindSectChain(flBranch,
                    yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch,
                    daiyunStem, daiyunBranch);

                // 小限太歲（男順女逆，以命宮地支第1歲起算）
                int suiAge = year - birthYear + 1; // 虛歲
                string xiaoXianBranch = LnCalcXiaoXianBranch(fallbackMingBranch, suiAge, gender);
                string xiaoXianGuard  = LnCalcXiaoXianGuard(xiaoXianBranch, fallbackMingBranch, suiAge, gender);

                var bestMonths    = monthlyDetails.Where(m => m.cross is "大吉" or "吉").OrderByDescending(m => m.bazi + m.ziwei).Take(3).Select(m => m.idx).ToList();
                var cautionMonths = monthlyDetails.Where(m => m.cross is "大凶" or "小凶").OrderBy(m => m.bazi + m.ziwei).Take(2).Select(m => m.idx).ToList();

                var baziTable = new {
                    pillars = new[] {
                        new { label = "年", stem = yStem, branch = yBranch, stemSS = yStemSS, naYin = LfPillarNaYin(yearP),  hiddenPairs = LfPillarHiddenPairs(yearP) },
                        new { label = "月", stem = mStem, branch = mBranch, stemSS = mStemSS, naYin = LfPillarNaYin(monthP), hiddenPairs = LfPillarHiddenPairs(monthP) },
                        new { label = "日", stem = dStem, branch = dBranch, stemSS = "元神",  naYin = LfPillarNaYin(dayP),   hiddenPairs = LfPillarHiddenPairs(dayP) },
                        new { label = "時", stem = hStem, branch = hBranch, stemSS = hStemSS, naYin = LfPillarNaYin(timeP),  hiddenPairs = LfPillarHiddenPairs(timeP) },
                    }
                };
                var scoredCycles = luckCycles.Select(lc => {
                    int sc = LfCalcLuckScore(lc.stem, lc.branch, pattern, yongShenElem, fuYiElem, jiShenElem,
                        dmElem, bodyPct > 50, tiaoHouElem, season, branches, chartStems4, dStem);
                    return new { lc.stem, lc.branch, lc.liuShen, lc.startAge, lc.endAge, score = sc, level = LfLuckLevel(sc) };
                }).ToList();

                var annualSummary = new {
                    year, stemBranch = flStem + flBranch,
                    currentDaiyun = daiyunStem + daiyunBranch,
                    daiyunAgeRange = $"{curLuck.startAge}-{curLuck.endAge}",
                    baziScore = flBaziScore, ziweiScore = flZiweiScore,
                    taisuiRelation, taisuiLevel, crossClass = flCrossClass,
                    bestMonths, cautionMonths
                };

                string report = LnBuildReport(
                    yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch,
                    yStemSS, mStemSS, hStemSS, yBranchSS, mBranchSS, dBranchSS, hBranchSS,
                    dmElem, wuXing, bodyPct, bodyLabel, season, seaLabel,
                    pattern, yongShenElem, fuYiElem, yongReason, jiShenElem,
                    luckCycles, daiyunStem, daiyunBranch, daiyunSS, daiyunBrSS,
                    curLuck.startAge, curLuck.endAge,
                    flStem, flBranch, flStemSS, flBranchSS, flBaziScore, flZiweiScore, flCrossClass,
                    year, flAge, gender, birthYear, shengXiaoBranch,
                    taisuiRelation, taisuiLevel, taisuiDesc, taisuiNeedAn, taisuiPos,
                    taisuiGen, taisuiBranchEvent,
                    taisuiGuard, xiaoXianGuard, blindSect,
                    hasZiwei, palaces, siHuaDescMap, monthlyDetails, bestMonths, cautionMonths,
                    branches, dStem);

                // 九星氣學加成（純 KB，流年版：命×運 + 命×流年）
                string lnNsSection = await LnNsBuildSection(
                    user.BirthYear ?? birthYear,
                    user.BirthMonth ?? 1,
                    user.BirthDay ?? 1,
                    user.BirthHour ?? 0,
                    user.BirthGender ?? gender,
                    year);
                if (!string.IsNullOrEmpty(lnNsSection)) report += lnNsSection;

                if (!lnIsAdmin) await RecordSubscriptionClaim(user.Id, lnSubId, "BOOK_LIUNIAN");
                await SaveUserReportAsync(user.Id, "liunian", $"{year} 流年命書", report,
                    new { year, birthYear = user.BirthYear, birthMonth = user.BirthMonth, birthDay = user.BirthDay, gender = user.BirthGender });
                return Ok(new { result = report, annualSummary, monthlyForecasts, baziTable, luckCycles = scoredCycles });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "流年命書失敗 User={User}", identity);
                return StatusCode(500, new { error = "流年命書生成失敗，請稍後再試", details = ex.Message });
            }
        }

        // === Ln (流年命書) Helper Methods ===

        private static (string stem, string branch) LnGetMonthStemBranch(string yearStem, int monthIdx)
        {
            string[] branches = { "寅","卯","辰","巳","午","未","申","酉","戌","亥","子","丑" };
            string[] stems    = { "甲","乙","丙","丁","戊","己","庚","辛","壬","癸" };
            int startIdx = yearStem switch { "甲" or "己" => 2, "乙" or "庚" => 4, "丙" or "辛" => 6, "丁" or "壬" => 8, _ => 0 };
            return (stems[(startIdx + monthIdx - 1) % 10], branches[monthIdx - 1]);
        }

        private static string LnGetMonthSeason(int monthIdx) => monthIdx switch
        {
            1 or 2 or 3 => "春",
            4 or 5 or 6 => "夏",
            7 or 8 or 9 => "秋",
            _            => "冬"
        };

        // 六十甲子太歲星君對照表
        private static readonly Dictionary<string, string> SixtyYearTaisuiGen = new() {
            {"甲子","金辦"},{"乙丑","陳材"},{"丙寅","耿章"},{"丁卯","沈興"},{"戊辰","趙達"},
            {"己巳","郭燦"},{"庚午","王濟"},{"辛未","李素"},{"壬申","劉旺"},{"癸酉","康志"},
            {"甲戌","施廣"},{"乙亥","任保"},{"丙子","郭嘉"},{"丁丑","汪文"},{"戊寅","魯先"},
            {"己卯","龍仲"},{"庚辰","董德"},{"辛巳","鄭但"},{"壬午","陸明"},{"癸未","魏仁"},
            {"甲申","方公"},{"乙酉","蔣崇"},{"丙戌","白敏"},{"丁亥","封濟"},{"戊子","鄒鐺"},
            {"己丑","傅儻"},{"庚寅","鄔桓"},{"辛卯","范寧"},{"壬辰","彭泰"},{"癸巳","徐單"},
            {"甲午","章詞"},{"乙未","楊仙"},{"丙申","管仲"},{"丁酉","唐杰"},{"戊戌","姜武"},
            {"己亥","謝壽"},{"庚子","虞起"},{"辛丑","楊信"},{"壬寅","賀諤"},{"癸卯","皮時"},
            {"甲辰","李誠"},{"乙巳","吳遂"},{"丙午","文哲"},{"丁未","繆丙"},{"戊申","徐浩"},
            {"己酉","程寶"},{"庚戌","倪秘"},{"辛亥","葉堅"},{"壬子","邱德"},{"癸丑","朱得"},
            {"甲寅","張朝"},{"乙卯","萬清"},{"丙辰","辛亞"},{"丁巳","楊彥"},{"戊午","黎卿"},
            {"己未","傅黨"},{"庚申","毛梓"},{"辛酉","石政"},{"壬戌","洪充"},{"癸亥","虞程"}
        };

        // 十二地支太歲年份民間事項（依流年地支，說明該年太歲特性）
        private static readonly Dictionary<string, string> BranchTaisuiEvent = new() {
            {"子","主水星動盪，智謀多變；財運流動不穩，宜守不宜進；防水液相關意外。"},
            {"丑","主土穩固，忍耐與積累；事業打基礎之年；財運緩中有穩，防情感波動。"},
            {"寅","主木衝勁，謀動求變；驛馬活絡，宜出行求職；防衝動決策造成損失。"},
            {"卯","主木桃花，文藝氣旺；感情易有動靜；宜學習進修、文書事務推進。"},
            {"辰","主土轉機，龍年多變；機會與挑戰並存；宜調整方向，防官非口舌。"},
            {"巳","主火謀略，多思多變；事業宜謀後而動；防情緒耗損、火氣過旺。"},
            {"午","主火奔波，財運活絡；積極求財旺；但防過度勞碌、衝動耗財。"},
            {"未","主土家庭，情感溫厚；家宅事宜多；宜調整人際關係，守財為主。"},
            {"申","主金聰明，驛馬奔波；變動多，適合轉換跑道；防多謀少決。"},
            {"酉","主金精細，口舌是非；注意人際關係；宜謹慎言行，防小人。"},
            {"戌","主土忠誠，變動收成；人生重大整合之年；宜總結過去，謀劃未來。"},
            {"亥","主水享樂，財氣流通；桃花人緣佳；防過度享樂導致財耗散。"}
        };

        // 十二歲君神煞表（索引0=位置1太歲 ... 索引11=位置12病符）
        private static readonly (string name, string alias, string luck, string desc, string advice)[] SuiJunTable = {
            ("太歲", "伏屍/劍鋒星", "需謹慎", "意外事故多，精神壓力大，萬事皆動", "少惹是非，少近官，閉門忍修"),
            ("太陽", "天空星",       "吉",     "貴人到位免災，出門辦事有人幫",     "積極把握貴人機緣，逢險可化解"),
            ("喪門", "地喪星",       "凶",     "孝服小災，親友有憂，防喪事來臨",   "關心長輩健康，若無內孝防外孝"),
            ("勾絞", "貫索星",       "平",     "平常度日，小口舌是非",             "男防口舌；女遇此有喜事機緣"),
            ("官符", "五鬼星",       "凶",     "官非是非，小人暗害，文書糾紛",     "謹言慎行，合同須慎，防官非訴訟"),
            ("死符", "小耗星",       "凶",     "損物破財，體力耗損，支出增加",     "謹防小人耗損，遇事請貴人協助"),
            ("歲破", "欄幹星",       "大凶",   "大耗破財，防水火意外，防拐騙病生", "勿大動作投資，守財為主，防病"),
            ("龍德", "天煞解厄",     "大吉",   "龍德護體，逢凶化吉，喜事可期",    "積極行事，婚育開創皆宜"),
            ("白虎", "飛廉星",       "凶",     "血光橫禍，孝服，防意外傷病開刀",  "不是內喪必外孝，防血光開刀"),
            ("福德", "捲舌星",       "吉",     "天德貴人助，求名求利順風",         "積極求財求名，廣結善緣"),
            ("天狗", "吊客星",       "凶",     "克兒克女，守喪守病，親人不平",     "關心家人健康，防親人有難"),
            ("病符", "吞陷星",       "凶",     "大病小病，天災人禍，辦事難",       "重視健康，勿拖延就醫，謹慎行事")
        };

        // 計算十二歲君位置（1-12），以 startBranch=1太歲，順數到 targetBranch
        private static int LnCalcSuiJunPos(string targetBranch, string startBranch)
        {
            string[] order = {"子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥"};
            int tIdx = Array.IndexOf(order, targetBranch);
            int sIdx = Array.IndexOf(order, startBranch);
            if (tIdx < 0 || sIdx < 0) return 0;
            return (tIdx - sIdx + 12) % 12 + 1;
        }

        private static (string relation, string level, string desc, string needAn, int suiJunPos) LnCalcTaisuiRelation(string shengXiaoBranch, string flowBranch)
        {
            // 生肖起算：以生肖年支為1（太歲），順數到流年地支，得神煞位置
            int pos = LnCalcSuiJunPos(flowBranch, shengXiaoBranch);
            string relation = pos switch {
                1  => "守太歲",
                7  => "沖太歲（歲破）",
                5  => "刑太歲（官符）",
                _  => SuiJunTable[pos - 1].name
            };
            var (sjName, sjAlias, sjLuck, sjDesc, sjAdvice) = SuiJunTable[pos - 1];
            string level = sjLuck;
            // 需安太歲判斷
            string needAn = pos switch {
                1 or 7         => "需安太歲（本命/沖太歲，壓力最大）",
                3 or 9 or 12   => "建議安太歲（喪門/白虎/病符）",
                5 or 6 or 11   => "建議祭解（官符/死符/天狗）",
                2 or 8 or 10   => "可安可不安（吉位，毋須特別安）",
                _              => "視個人情況，可安保平安"
            };
            string desc = $"歲君【{pos}.{sjName}（{sjAlias}）】：{sjDesc}。{sjAdvice}。";
            return (relation, level, desc, needAn, pos);
        }

        private static string LnCalcMingGongGuard(string flowBranch, JsonElement palaces, string fallbackMingBranch = "")
        {
            // 十二歲君以八字命宮為準（月支+時支公式），不使用紫微命宮地支
            string mingBranch = fallbackMingBranch;
            if (string.IsNullOrEmpty(mingBranch)) return "";

            // 命宮起算十二歲君（情境二：以命宮地支為1太歲，順數到流年地支）
            int mgPos = LnCalcSuiJunPos(flowBranch, mingBranch);
            var (mgSjName, mgSjAlias, mgSjLuck, mgSjDesc, mgSjAdvice) = SuiJunTable[mgPos - 1];

            var sbMg = new System.Text.StringBuilder();
            sbMg.AppendLine($"命宮地支：{mingBranch}宮");
            sbMg.AppendLine($"逢 第 {mgPos} 位歲君：【{mgSjName}（{mgSjAlias}）】");
            sbMg.AppendLine($"  代表個人身心：{mgSjDesc}");
            sbMg.Append($"  建議：{mgSjAdvice}");

            // 太歲與命宮的位置關係（純地支距離判斷，無紫微星曜）
            int dist = ((Array.IndexOf(new[]{"子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥"}, flowBranch)
                       - Array.IndexOf(new[]{"子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥"}, mingBranch)) + 12) % 12;
            string posDesc = dist switch {
                0       => $"今年太歲正臨命宮（{mingBranch}宮）：守宮太歲，人生大動，積極迎接變化，宜破舊立新。",
                6       => $"今年太歲沖命宮（太歲在{flowBranch}宮，命宮在{mingBranch}宮，六沖）：正沖命宮，六親宮位均受震動，宜低調防守，勿輕舉妄動。",
                1 or 11 => $"今年太歲臨命宮鄰宮（{flowBranch}宮，距命宮 {dist} 位）：文書長輩事宜活躍，感受到外部壓力，官府文件需留意。",
                2 or 10 => $"今年太歲臨命宮（{flowBranch}宮，距 {dist} 位）：合夥人際事宜較多，謹慎合約，防口舌。",
                3 or 9  => $"今年太歲臨命宮（{flowBranch}宮，距 {dist} 位）：感情婚姻有重大進展或波動，感情面向受動。",
                4 or 8  => $"今年太歲臨命宮（{flowBranch}宮，距 {dist} 位）：桃花緣分活躍，子女事宜有動靜。",
                5 or 7  => $"今年太歲臨命宮（{flowBranch}宮，距 {dist} 位）：財運明顯變動，宜把握機遇。",
                _       => $"今年太歲行至{flowBranch}宮，距命宮 {dist} 位，整體平穩行事。"
            };
            sbMg.AppendLine();
            sbMg.Append(posDesc);
            return sbMg.ToString().Trim();
        }

        // 盲派串宮壓運：12神（串宮版），以流年地支起太歲，順數12支
        private static readonly string[] BlindSect12Names = {
            "太歲","青龍","喪門","六合","官符","小耗","大耗","朱雀","白虎","貴神","吊客","病符"
        };
        private static readonly string[] BlindSect12Luck = {
            "中性","吉","凶","吉","中性","小凶","大凶","大凶","大凶","吉","小凶","小凶"
        };
        private static readonly string[] BlindSect12Desc = {
            "主變動壓頂，吉凶取決於配合，見凶則凶",
            "主光明喜事，貴人相助，事業財運順暢",
            "主孝服破財，災疾，親友有憂事",
            "主喜慶合作，人丁添旺，感情合婚順遂",
            "主官非口舌，文書糾紛，財運有爭議",
            "主財物耗散，奔波不順，小破財損耗",
            "主大破財變動，官司竊騙，凶事較重",
            "主口舌是非，背黑鍋，小人橫行作祟",
            "主血光意外，病災官司，孝服傷亡",
            "主貴人財喜，官運亨通，福祿安泰",
            "主小人受損，驚嚇不安，親人不平",
            "主身體病痛，憂愁纏身，辦事多阻礙"
        };
        // 天干通地支（盲派規則）
        private static readonly Dictionary<string, string> BlindSectStemToBranch = new()
        {
            {"甲","寅"},{"乙","卯"},{"丙","巳"},{"丁","午"},{"戊","戌"},
            {"己","丑"},{"庚","申"},{"辛","酉"},{"壬","亥"},{"癸","子"}
        };

        private static readonly Dictionary<string, string> LnStarAbbrToFullMap = new()
        {
            {"廉","廉貞"},{"破","破軍"},{"武","武曲"},{"陽","太陽"},
            {"機","天機"},{"梁","天梁"},{"紫","紫微"},{"陰","太陰"},
            {"同","天同"},{"巨","巨門"},{"貪","貪狼"},{"相","天相"},
            {"殺","七殺"},{"府","天府"},{"昌","文昌"},{"曲","文曲"},
            {"輔","左輔"},{"弼","右弼"}
        };

        private static string LnStarAbbrFull(string abbr) =>
            LnStarAbbrToFullMap.TryGetValue(abbr, out var full) ? full : abbr;

        private static string LnCalcBlindSectChain(
            string flowBranch,
            string yStem, string yBranch,
            string mStem, string mBranch,
            string dStem, string dBranch,
            string hStem, string hBranch,
            string daiyunStem, string daiyunBranch)
        {
            string[] brStd = {"子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥"};
            int flowIdx = Array.IndexOf(brStd, flowBranch);
            if (flowIdx < 0) return "";

            int GetBrPos(string br) {
                int i = Array.IndexOf(brStd, br);
                return i < 0 ? 0 : (i - flowIdx + 12) % 12 + 1;
            }
            int GetStPos(string st) {
                if (!BlindSectStemToBranch.TryGetValue(st, out var tb)) return 0;
                return GetBrPos(tb);
            }
            string ShenName(int p) => p >= 1 && p <= 12 ? BlindSect12Names[p-1] : "";
            string ShenLuck(int p) => p >= 1 && p <= 12 ? BlindSect12Luck[p-1] : "";
            string ShenDesc(int p) => p >= 1 && p <= 12 ? BlindSect12Desc[p-1] : "";
            bool IsGood(int p)     => p >= 1 && p <= 12 && BlindSect12Luck[p-1] is "吉";

            var sb2 = new StringBuilder();

            // 大運（全年根基）
            if (!string.IsNullOrEmpty(daiyunBranch))
            {
                int dyBrP  = GetBrPos(daiyunBranch);
                int dyStP  = !string.IsNullOrEmpty(daiyunStem) ? GetStPos(daiyunStem) : 0;
                sb2.AppendLine($"大運 {daiyunStem}{daiyunBranch}（全年根基）：");
                if (dyStP > 0)
                {
                    string tb = BlindSectStemToBranch.GetValueOrDefault(daiyunStem, "");
                    sb2.AppendLine($"  干 {daiyunStem}通{tb} → 【{ShenName(dyStP)}】（{ShenLuck(dyStP)}）{ShenDesc(dyStP)}");
                }
                if (dyBrP > 0)
                    sb2.AppendLine($"  支 {daiyunBranch} → 【{ShenName(dyBrP)}】（{ShenLuck(dyBrP)}）{ShenDesc(dyBrP)}");
                bool dyBad = (dyStP > 0 && !IsGood(dyStP) && BlindSect12Luck[dyStP-1] != "中性")
                          || (dyBrP > 0 && !IsGood(dyBrP) && BlindSect12Luck[dyBrP-1] != "中性");
                bool dyGood = (dyStP <= 0 || IsGood(dyStP)) && (dyBrP <= 0 || IsGood(dyBrP));
                sb2.AppendLine(dyBad  ? "  根基評斷：大運見凶，全年根基受壓，各季吉星作用減弱，宜謹慎守成。"
                             : dyGood ? "  根基評斷：大運見吉，全年根基穩健，吉星助力，可積極有為。"
                             :          "  根基評斷：大運吉凶互見，全年需視各季應事輕重行事。");
                sb2.AppendLine();
            }

            // 四柱分析（年=春/月=夏/日=秋/時=冬）
            (string pName, string season, string subj, string st, string br)[] pillars = {
                ("年柱", "春", "父母祖業", yStem, yBranch),
                ("月柱", "夏", "事業財富", mStem, mBranch),
                ("日柱", "秋", "自身配偶", dStem, dBranch),
                ("時柱", "冬", "子女晚輩", hStem, hBranch),
            };
            foreach (var (pName, season, subj, st, br) in pillars)
            {
                int brP  = GetBrPos(br);
                int stP  = GetStPos(st);
                if (brP <= 0) continue;
                string tb = BlindSectStemToBranch.GetValueOrDefault(st, "");
                sb2.AppendLine($"{pName}（{season}季/{subj}）{st}{br}：");
                if (stP > 0)
                    sb2.AppendLine($"  干 {st}通{tb} → 【{ShenName(stP)}】（{ShenLuck(stP)}）{ShenDesc(stP)}");
                sb2.AppendLine($"  支 {br} → 【{ShenName(brP)}】（{ShenLuck(brP)}）{ShenDesc(brP)}");
                bool bad  = (stP > 0 && BlindSect12Luck[stP-1] is "凶" or "大凶")
                         || BlindSect12Luck[brP-1] is "凶" or "大凶";
                bool good = (stP <= 0 || BlindSect12Luck[stP-1] is "吉")
                         && BlindSect12Luck[brP-1] is "吉";
                string eval = bad  ? $"此季凶星壓{subj}，{season}季需特別留意" :
                              good ? $"此季吉星照{subj}，{season}季可積極把握" :
                                     $"此季吉凶參半，{season}季平穩行事";
                sb2.AppendLine($"  評斷：{eval}");
            }
            return sb2.ToString().Trim();
        }

        // 命宮計算（純八字公式，與紫微無關）
        // 地支序：寅=1,卯=2,辰=3,巳=4,午=5,未=6,申=7,酉=8,戌=9,亥=10,子=11,丑=12
        // 公式：sum=月支數+時支數; sum<14 → 14-sum; sum>=14 → 26-sum
        // 中氣修正：若出生日已過當月中氣（雨水/春分/穀雨等12個），月支+1
        // 地支標準順序（子=0...亥=11），供小限順逆行使用
        private static readonly string[] LnBranchStd = { "子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥" };

        // 計算小限地支：命宮地支第1歲，男順女逆
        private static string LnCalcXiaoXianBranch(string mingBranch, int suiAge, int gender)
        {
            if (string.IsNullOrEmpty(mingBranch) || suiAge < 1) return "";
            int mingIdx = Array.IndexOf(LnBranchStd, mingBranch);
            if (mingIdx < 0) return "";
            int steps = (suiAge - 1) % 12;
            int idx = gender == 1
                ? (mingIdx + steps) % 12       // 男順行
                : (mingIdx - steps + 120) % 12; // 女逆行
            return LnBranchStd[idx];
        }

        // 小限太歲十二歲君報告（以命宮地支為第1太歲，小限地支起算）
        private static string LnCalcXiaoXianGuard(string xiaoXianBranch, string mingBranch, int suiAge, int gender)
        {
            if (string.IsNullOrEmpty(xiaoXianBranch) || string.IsNullOrEmpty(mingBranch)) return "";
            int pos = LnCalcSuiJunPos(xiaoXianBranch, mingBranch);
            if (pos < 1 || pos > 12) return "";
            var (sjName, sjAlias, sjLuck, sjDesc, sjAdvice) = SuiJunTable[pos - 1];
            string dir = gender == 1 ? "男順行" : "女逆行";
            var sb = new StringBuilder();
            sb.AppendLine($"虛歲 {suiAge} 歲，小限落在：{xiaoXianBranch}宮");
            sb.AppendLine($"小限第 {pos} 位歲君：【{sjName}（{sjAlias}）】");
            sb.AppendLine($"  代表個人身心：{sjDesc}");
            sb.Append($"  建議：{sjAdvice}");
            return sb.ToString().Trim();
        }

        // 命宮地支序（寅=1,卯=2...丑=12），用於命宮公式
        private static readonly string[] LnMgBranchOrder = { "寅","卯","辰","巳","午","未","申","酉","戌","亥","子","丑" };
        private const string LnZhongQiList = "'雨水','春分','穀雨','小滿','夏至','大暑','處暑','秋分','霜降','小雪','冬至','大寒'";

        private static string LnCalcMingBranchFromBazi(string mBranch, string hBranch)
        {
            // 八字月柱地支已確定月份，直接套命宮公式，不再做中氣修正
            // 地支序：寅=1,卯=2,...,亥=10,子=11,丑=12
            int mIdx = Array.IndexOf(LnMgBranchOrder, mBranch);
            int hIdx = Array.IndexOf(LnMgBranchOrder, hBranch);
            if (mIdx < 0 || hIdx < 0) return "";
            int mNum = mIdx + 1;
            int hNum = hIdx + 1;
            int sum  = mNum + hNum;
            int mgNum = sum < 14 ? 14 - sum : 26 - sum;
            if (mgNum < 1 || mgNum > 12) return "";
            return LnMgBranchOrder[mgNum - 1];
        }

        private static string LnGetMonthFlowStar(int monthIdx, string flowBranch, JsonElement palaces)
        {
            if (palaces.ValueKind != JsonValueKind.Array) return "";
            string[] branchOrder = { "子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥" };
            int startIdx = Array.IndexOf(branchOrder, flowBranch);
            if (startIdx < 0) return "";
            string targetBranch = branchOrder[(startIdx + monthIdx - 1) % 12];
            string targetPalace = KbGetPalaceByBranch(palaces, targetBranch);
            if (string.IsNullOrEmpty(targetPalace)) return $"流月行至{targetBranch}位";
            string stars = KbGetPalaceStars(palaces, targetPalace);
            return string.IsNullOrEmpty(stars)
                ? $"流月行至{targetPalace}（空宮）"
                : $"流月行至{targetPalace}（{stars}）";
        }

        private static string LnMonthTip(string crossClass, string season, string stemSS) => crossClass switch
        {
            "大吉" => $"{season}季此月大吉，宜積極進取，把握商機，{stemSS}氣場最旺。",
            "吉"   => $"{season}季整體向好，宜穩健推進計劃，{stemSS}助力明顯。",
            "平"   => $"{season}季平穩守成，避免冒進，靜待時機。",
            "小凶" => $"{season}季需謹慎，防是非耗損，{stemSS}壓力較大。",
            "大凶" => $"{season}季壓力最重，宜低調守成，避免重大決策。",
            _      => "平穩行事，量力而為。"
        };

        private static string LnBuildReport(
            string yStem, string yBranch, string mStem, string mBranch,
            string dStem, string dBranch, string hStem, string hBranch,
            string yStemSS, string mStemSS, string hStemSS,
            string yBranchSS, string mBranchSS, string dBranchSS, string hBranchSS,
            string dmElem, Dictionary<string, double> wuXing, double bodyPct, string bodyLabel,
            string season, string seaLabel, string pattern,
            string yongShenElem, string fuYiElem, string yongReason, string jiShenElem,
            List<(string stem, string branch, string liuShen, int startAge, int endAge)> luckCycles,
            string daiyunStem, string daiyunBranch, string daiyunSS, string daiyunBrSS,
            int daiyunStartAge, int daiyunEndAge,
            string flStem, string flBranch, string flStemSS, string flBranchSS,
            int flBaziScore, int flZiweiScore, string flCrossClass,
            int year, int flAge, int gender, int birthYear, string shengXiaoBranch,
            string taisuiRelation, string taisuiLevel, string taisuiDesc, string taisuiNeedAn, int taisuiPos,
            string taisuiGen, string taisuiBranchEvent,
            string taisuiGuard, string xiaoXianGuard, string blindSect,
            bool hasZiwei, JsonElement palaces,
            Dictionary<string, Dictionary<string, (string palace, string desc)>> siHuaDescMap,
            List<(int idx, string mStemM, string mBranchM, string mSeason, int bazi, int ziwei, string cross, string flowStar, string tip)> monthlyDetails,
            List<int> bestMonths, List<int> cautionMonths,
            string[] branches, string dStemRef)
        {
            var sb = new StringBuilder();
            string genderText = gender == 1 ? "男（乾造）" : "女（坤造）";
            string SS(string ss) => string.IsNullOrEmpty(ss) ? "" : $"（{ss}）";
            string wx = $"木{wuXing["木"]:F0}% 火{wuXing["火"]:F0}% 土{wuXing["土"]:F0}% 金{wuXing["金"]:F0}% 水{wuXing["水"]:F0}%";
            string[] monthNames = { "寅月(2月)","卯月(3月)","辰月(4月)","巳月(5月)","午月(6月)","未月(7月)","申月(8月)","酉月(9月)","戌月(10月)","亥月(11月)","子月(12月)","丑月(1月)" };
            string tuneElem   = season == "冬" ? "火" : season == "夏" ? "水" : "";
            string jiYongElem = LfElemOvercomeBy.GetValueOrDefault(yongShenElem, "");

            sb.AppendLine("=================================================================");
            sb.AppendLine("                         流 年 命 書");
            sb.AppendLine("=================================================================");
            sb.AppendLine();

            // 人生指南目錄
            sb.AppendLine("                       人  生  指  南");
            sb.AppendLine("-----------------------------------------------------------------");
            sb.AppendLine("  命主資料與流年概況");
            sb.AppendLine("  格局用神與流年八字分析");
            sb.AppendLine("  流年小限空間與時間影響");
            sb.AppendLine("  春夏秋冬四季論斷");
            sb.AppendLine("  逐月分析（月建喜忌・紫微宮位）");
            sb.AppendLine("  趨吉避凶全年建議");
            sb.AppendLine("  九星氣學流年加成");
            sb.AppendLine("  人生警示事項（先天體質）");
            sb.AppendLine("-----------------------------------------------------------------");
            sb.AppendLine();

            // Ch.1 命主資料 + 流年概況
            sb.AppendLine("【第一章：命主資料與流年概況】");
            sb.AppendLine($"性別：{genderText}  出生年：{birthYear} 年（生肖屬{shengXiaoBranch}）");
            sb.AppendLine($"四柱：{yStem}{yBranch} {mStem}{mBranch} {dStem}{dBranch} {hStem}{hBranch}");
            sb.AppendLine($"十神：年干{SS(yStemSS)} 年支{SS(yBranchSS)} 月干{SS(mStemSS)} 月支{SS(mBranchSS)} 時干{SS(hStemSS)} 時支{SS(hBranchSS)}");
            sb.AppendLine($"日主：{dStem}（{dmElem}）  格局：{pattern}  日主{bodyLabel}（{bodyPct:F0}%）");
            sb.AppendLine($"用神：{yongShenElem}  忌神：{jiShenElem}  五行：{wx}");
            sb.AppendLine();
            sb.AppendLine($"分析年份：{year} 年  流年：{flStem}{flBranch}（天干{flStemSS}·地支{flBranchSS}）  年齡：{flAge} 歲");
            sb.AppendLine($"流年整體：八字 {flBaziScore} 分  紫微 {flZiweiScore} 分  綜合：【{flCrossClass}】");
            if (!string.IsNullOrEmpty(daiyunStem))
                sb.AppendLine($"當前大運：{daiyunStem}{daiyunBranch}（天干{daiyunSS}·地支{daiyunBrSS}），{daiyunStartAge}-{daiyunEndAge} 歲");
            sb.AppendLine($"生肖太歲：{shengXiaoBranch} 生肖，{year} 年太歲{flBranch}，關係【{taisuiRelation}】（{taisuiLevel}）");
            sb.AppendLine();

            // Ch.2 格局用神 + 流年八字分析
            sb.AppendLine("【第二章：格局用神與流年八字分析】");
            sb.AppendLine($"格局：【{pattern}】  用神：【{yongShenElem}】（{yongReason}）");
            sb.AppendLine($"喜用天干：{LfElemStems(yongShenElem)}  喜用地支：{LfElemBranches(yongShenElem)}");
            if (fuYiElem != yongShenElem)
                sb.AppendLine($"輔助喜神：【{fuYiElem}】");
            sb.AppendLine($"大忌(X)：{jiShenElem}  天干 {LfElemStems(jiShenElem)}  地支 {LfElemBranches(jiShenElem)}");
            if (!string.IsNullOrEmpty(jiYongElem) && jiYongElem != jiShenElem)
                sb.AppendLine($"次忌(△忌)：{jiYongElem}（克用神{yongShenElem}）");
            sb.AppendLine();
            sb.AppendLine(LfBuildYongJiTable(yongShenElem, fuYiElem, jiShenElem, tuneElem, dStemRef, branches));
            sb.AppendLine();

            // 流年天干地支喜忌
            string flStemElem = KbStemToElement(flStem);
            string flBrMs2    = LfBranchHiddenRatio.TryGetValue(flBranch, out var flBrH2) && flBrH2.Count > 0 ? flBrH2[0].stem : "";
            string flBrElem2  = !string.IsNullOrEmpty(flBrMs2) ? KbStemToElement(flBrMs2) : "";
            string stemCls    = (flStemElem == yongShenElem || flStemElem == fuYiElem) ? "○（喜用）" : flStemElem == jiShenElem ? "X（大忌）" : "△（中性）";
            string brCls      = string.IsNullOrEmpty(flBrElem2) ? "△（中性）" : (flBrElem2 == yongShenElem || flBrElem2 == fuYiElem) ? "○（喜用）" : flBrElem2 == jiShenElem ? "X（大忌）" : "△（中性）";
            sb.AppendLine($"流年天干 {flStem}（{flStemElem}·{flStemSS}）：{stemCls}");
            if (!string.IsNullOrEmpty(flBrElem2))
                sb.AppendLine($"流年地支 {flBranch}（{flBrElem2}·{flBranchSS}）：{brCls}");
            string brEvents = LfBranchEvents(flBranch, branches);
            if (!string.IsNullOrEmpty(brEvents)) sb.AppendLine($"歲君互動：{brEvents}");

            if (!string.IsNullOrEmpty(daiyunStem))
            {
                string dyElem  = KbStemToElement(daiyunStem);
                bool dyHelps   = dyElem == yongShenElem || dyElem == fuYiElem;
                bool flHelps   = flStemElem == yongShenElem || flStemElem == fuYiElem;
                string inter   = (dyHelps && flHelps) ? "大運流年雙喜，力量倍增，宜積極進取。"
                               : (!dyHelps && !flHelps) ? "大運流年均不利用神，壓力較重，宜謹慎守成。"
                               : dyHelps ? "大運喜用，流年稍有阻礙，整體仍向好。"
                               : "流年喜用，大運壓力尚存，需把握流年窗口期積極行動。";
                sb.AppendLine($"大運流年互動：{daiyunStem}{daiyunBranch}運 + {flStem}{flBranch}年 → {inter}");
            }
            sb.AppendLine();
            sb.AppendLine($"【流年綜合論斷】{DyCrossDesc(flCrossClass, flStemSS, flBranchSS, flBaziScore, flZiweiScore)}");
            sb.AppendLine();

            // Ch.3 民間五術加成
            sb.AppendLine("【第三章：流年小限空間與時間影響】");
            sb.AppendLine();
            sb.AppendLine("【3.1 流年歲君】");
            if (string.IsNullOrEmpty(taisuiGuard))
                sb.AppendLine("（命宮資料不足，無法計算流年歲君）");
            else
                sb.AppendLine(taisuiGuard);
            sb.AppendLine();
            sb.AppendLine("【3.1b 小限歲君】");
            if (string.IsNullOrEmpty(xiaoXianGuard))
                sb.AppendLine("（命宮資料不足，無法計算小限歲君）");
            else
                sb.AppendLine(xiaoXianGuard);
            sb.AppendLine();
            sb.AppendLine("【3.2 生肖十二歲君】");
            string genLabel = string.IsNullOrEmpty(taisuiGen) ? "" : $"值年太歲星君：{taisuiGen}大將軍";
            sb.AppendLine($"{year} 年（{flBranch}年）{genLabel}");
            if (!string.IsNullOrEmpty(taisuiBranchEvent))
                sb.AppendLine($"  特性：{taisuiBranchEvent}");
            sb.AppendLine();
            var (sjName1, sjAlias1, sjLuck1, sjDesc1, sjAdvice1) = SuiJunTable[taisuiPos - 1];
            sb.AppendLine($"【生肖起算（外在環境）】：生肖屬{shengXiaoBranch}，{year}年遇");
            sb.AppendLine($"  歲君：【{taisuiPos}.{sjName1}（{sjAlias1}）】");
            sb.AppendLine($"  外在環境：{sjDesc1}");
            sb.AppendLine($"  建議：{sjAdvice1}");
            sb.AppendLine($"  安太歲：{taisuiNeedAn}");
            sb.AppendLine();
            sb.AppendLine("【3.3 流星壓運】");
            sb.AppendLine(blindSect);
            sb.AppendLine();
            sb.AppendLine("【3.4 流年紫微四化】");
            if (hasZiwei && YearStemSiHuaMap.TryGetValue(flStem, out var siHua))
            {
                string[] shTypes = { "化祿", "化權", "化科", "化忌" };
                string[] shStars = { siHua.lu, siHua.quan, siHua.ke, siHua.ji };
                if (siHuaDescMap.TryGetValue(flStem, out var flMap))
                {
                    for (int si = 0; si < shTypes.Length; si++)
                    {
                        var (pal, desc) = flMap.GetValueOrDefault(shTypes[si], ("", ""));
                        string palLabel = string.IsNullOrEmpty(pal) ? "（命盤未含此星）" : $"入{pal}";
                        string descText = string.IsNullOrEmpty(desc) ? "" : $"：{desc}";
                        sb.AppendLine($"  {shTypes[si]}（{shStars[si]}星）{palLabel}{descText}");
                    }
                }
                else
                {
                    for (int si = 0; si < shTypes.Length; si++)
                    {
                        string pal = KbGetSiHuaPalace(flStem, shTypes[si], palaces);
                        sb.AppendLine($"  {shTypes[si]}（{shStars[si]}星）{(string.IsNullOrEmpty(pal) ? "（命盤未含此星）" : "入" + pal)}");
                    }
                }
            }
            else sb.AppendLine("（無紫微命盤資料，流年四化僅列星名參考）");
            sb.AppendLine();

            // Ch.4 春夏秋冬四季
            sb.AppendLine("【第四章：春夏秋冬四季論斷】");
            sb.AppendLine();
            var seasonGroups = new[] {
                ("春", new[] {1,2,3}, "寅月(2月)~辰月(4月)", "木"),
                ("夏", new[] {4,5,6}, "巳月(5月)~未月(7月)", "火"),
                ("秋", new[] {7,8,9}, "申月(8月)~戌月(10月)", "金"),
                ("冬", new[] {10,11,12}, "亥月(11月)~丑月(1月)", "水"),
            };
            foreach (var (sName, mIdxes, sRange, sWang) in seasonGroups)
            {
                var sMths = monthlyDetails.Where(m => mIdxes.Contains(m.idx)).ToList();
                double sAvgBazi  = sMths.Average(m => (double)m.bazi);
                double sAvgZiwei = sMths.Average(m => (double)m.ziwei);
                string sCross    = DyCrossClass((int)sAvgBazi, (int)sAvgZiwei);
                var sBest = sMths.Where(m => m.cross is "大吉" or "吉").Select(m => monthNames[m.idx-1]).ToList();
                var sCaut = sMths.Where(m => m.cross is "大凶" or "小凶").Select(m => monthNames[m.idx-1]).ToList();
                string sYongMatch = (sWang == yongShenElem || sWang == fuYiElem)
                    ? $"季節{sWang}旺，與用神{yongShenElem}相輔，整體得力。"
                    : sWang == jiShenElem ? $"季節{sWang}旺，忌神得勢，壓力偏大，需謹慎。"
                    : $"季節{sWang}旺，對用神{yongShenElem}影響中性。";
                sb.AppendLine($"【{sName}季】{sRange}（{sWang}旺）  評分：八字{sAvgBazi:F0}·紫微{sAvgZiwei:F0}·綜合【{sCross}】");
                sb.AppendLine($"  八字面向：{sYongMatch}");
                if (sBest.Count > 0) sb.AppendLine($"  本季佳月：{string.Join("、", sBest)}");
                if (sCaut.Count > 0) sb.AppendLine($"  本季謹慎：{string.Join("、", sCaut)}");
                sb.AppendLine();
            }

            // Ch.5 逐月分析
            sb.AppendLine("【第五章：逐月分析（月建喜忌·流月星曜·紫微宮位）】");
            sb.AppendLine();
            foreach (var m in monthlyDetails)
            {
                string mStemSS2   = LfStemShiShen(m.mStemM, dStemRef);
                string mBrMs2     = LfBranchHiddenRatio.TryGetValue(m.mBranchM, out var mBrH2) && mBrH2.Count > 0 ? mBrH2[0].stem : "";
                string mBrSSM     = !string.IsNullOrEmpty(mBrMs2) ? LfStemShiShen(mBrMs2, dStemRef) : "";
                string mStemElemM = KbStemToElement(m.mStemM);
                string mBrElemM   = !string.IsNullOrEmpty(mBrMs2) ? KbStemToElement(mBrMs2) : "";
                string mStemCls   = (mStemElemM == yongShenElem || mStemElemM == fuYiElem) ? "○喜用" : mStemElemM == jiShenElem ? "X忌" : "△中性";
                string mBrCls     = string.IsNullOrEmpty(mBrElemM) ? "" : (mBrElemM == yongShenElem || mBrElemM == fuYiElem) ? "○喜用" : mBrElemM == jiShenElem ? "X忌" : "△中性";
                sb.AppendLine($"【{monthNames[m.idx-1]}】{m.mStemM}{m.mBranchM}（{m.mSeason}季）  綜合：【{m.cross}】  八字{m.bazi}·紫微{m.ziwei}");
                sb.AppendLine($"  月建喜忌：天干{m.mStemM}（{mStemElemM}·{mStemSS2}）{mStemCls}{(!string.IsNullOrEmpty(mBrCls) ? $"  地支{m.mBranchM}（{mBrElemM}·{mBrSSM}）{mBrCls}" : "")}");
                if (hasZiwei && YearStemSiHuaMap.TryGetValue(m.mStemM, out var mSiHua))
                {
                    string[] branchOrd = {"子","丑","寅","卯","辰","巳","午","未","申","酉","戌","亥"};
                    string[] palaceOrd = {"命宮","兄弟宮","夫妻宮","子女宮","財帛宮","疾厄宮","遷移宮","交友宮","官祿宮","田宅宮","福德宮","父母宮"};
                    int mingIdxM = Array.IndexOf(branchOrd, m.mBranchM);
                    if (mingIdxM >= 0)
                    {
                        var siHuaEntries = new[] {
                            ("月化祿", mSiHua.lu),
                            ("月化權", mSiHua.quan),
                            ("月化科", mSiHua.ke),
                            ("月化忌", mSiHua.ji)
                        };
                        foreach (var (siHuaLabel, starAbbr) in siHuaEntries)
                        {
                            if (string.IsNullOrEmpty(starAbbr)) continue;
                            string starFullName = LnStarAbbrFull(starAbbr);
                            string palaceNameZiwei = KbFindPalaceByStarAbbr(palaces, starAbbr);
                            if (string.IsNullOrEmpty(palaceNameZiwei)) continue;
                            string starBranch = KbGetPalaceBranch(palaces, palaceNameZiwei);
                            int starBranchIdx = Array.IndexOf(branchOrd, starBranch);
                            if (starBranchIdx < 0) continue;
                            int offset = (mingIdxM - starBranchIdx + 12) % 12;
                            string monthPalace = palaceOrd[offset];
                            sb.AppendLine($"  {siHuaLabel}（{starFullName}）入{monthPalace}");
                        }
                    }
                }
                sb.AppendLine($"  本月提示：{m.tip}");
                sb.AppendLine();
            }

            // Ch.6 趨吉避凶
            sb.AppendLine("【第六章：趨吉避凶全年建議】");
            var allGood = monthlyDetails.Where(m => m.cross is "大吉" or "吉").OrderByDescending(m => m.bazi + m.ziwei).ToList();
            var allBad  = monthlyDetails.Where(m => m.cross is "大凶" or "小凶").OrderBy(m => m.bazi + m.ziwei).ToList();
            if (allGood.Count > 0)
            {
                sb.AppendLine("最佳把握月份：");
                sb.AppendLine($"  {string.Join("、", allGood.Take(4).Select(m => $"{monthNames[m.idx-1]}（{m.cross}·{m.mStemM}{m.mBranchM}）"))}");
                sb.AppendLine("  此類月份宜積極展開事業、感情、投資理財佈局。");
            }
            if (allBad.Count > 0)
            {
                sb.AppendLine("需謹慎月份：");
                sb.AppendLine($"  {string.Join("、", allBad.Take(3).Select(m => $"{monthNames[m.idx-1]}（{m.cross}·{m.mStemM}{m.mBranchM}）"))}");
                sb.AppendLine("  此類月份宜低調保守，避免重大決策，注意健康與財務風險。");
            }
            sb.AppendLine();
            sb.AppendLine("具體行動建議：");
            sb.AppendLine($"  1. 喜用方位：{LfElemDir(yongShenElem)}  喜用色彩：{LfElemColor(yongShenElem)}");
            sb.AppendLine($"  2. 忌諱方向：{LfElemDir(jiShenElem)}，減少{jiShenElem}屬性人事物接觸");
            sb.AppendLine($"  3. 事業宜從事：{LfElemCareer(yongShenElem)}");
            if (taisuiLevel is "凶" or "小凶")
                sb.AppendLine($"  4. 太歲化解：{year} 年{taisuiRelation}，建議安太歲、多行善事、順應變動而非強行抵抗。");
            else if (taisuiLevel is "吉" or "變動")
                sb.AppendLine($"  4. 太歲加持：{year} 年{taisuiRelation}，善用此年順勢而為，積極展開重要計劃。");
            sb.AppendLine($"  5. {DyCrossDesc(flCrossClass, flStemSS, flBranchSS, flBaziScore, flZiweiScore)}");
            sb.AppendLine();
            // === 人生警示事項 ===
            sb.AppendLine("【人生警示事項（先天體質）】");
            sb.AppendLine();
            sb.AppendLine("▍ 小人防範");
            sb.AppendLine(LfXiaoRenAnalysis(yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch, jiShenElem, dmElem));
            sb.AppendLine();
            sb.AppendLine("▍ 官司文書風險");
            sb.AppendLine(LfGuanSiAnalysis(yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch, jiShenElem, dmElem, bodyPct));
            sb.AppendLine();
            sb.AppendLine("▍ 車關時機");
            {
                string cheGuanBase = LfCheGuanAnalysis(yBranch, mBranch, dBranch, hBranch, jiShenElem, dmElem);
                sb.AppendLine(cheGuanBase);
                // 本年是否引動車關
                string guanShaElemLn = LfElemOvercomeBy.GetValueOrDefault(dmElem, "");
                var siYiMaLn = new HashSet<string> { "寅", "申", "巳", "亥" };
                var riskBranchesLn = new[] { yBranch, mBranch, dBranch, hBranch }
                    .Where(b => siYiMaLn.Contains(b) && LfBranchElem.GetValueOrDefault(b, "") == guanShaElemLn)
                    .ToList();
                if (riskBranchesLn.Count > 0)
                {
                    bool yearTriggered = riskBranchesLn.Any(rb =>
                        LfChong.Contains(rb + flBranch) || LfChong.Contains(flBranch + rb));
                    if (yearTriggered)
                        sb.AppendLine($"【本年車關提醒】{year} 年流年地支{flBranch}沖動命盤驛馬官殺，本年車關風險提升，請加強交通安全意識。");
                }
            }
            sb.AppendLine();
            sb.AppendLine("▍ 海外發展");
            sb.AppendLine(LfHaiWaiAnalysis(yBranch, mBranch, dBranch, hBranch, yongShenElem, jiShenElem, dmElem, hasZiwei, palaces));
            sb.AppendLine();
            sb.AppendLine("▍ 天乙貴人方向");
            {
                var tianYiMapLn = new Dictionary<string, string>
                {
                    {"甲","丑未"},{"戊","丑未"},{"庚","丑未"},
                    {"乙","子申"},{"己","子申"},
                    {"丙","亥酉"},{"丁","亥酉"},
                    {"壬","卯巳"},{"癸","卯巳"},
                    {"辛","午寅"}
                };
                string tianYiBranchesLn = tianYiMapLn.GetValueOrDefault(dStem, "");
                sb.AppendLine($"{dStem} 日主，天乙貴人在：{tianYiBranchesLn}（見此地支方位或流年行此地支，貴人助力最強）");
            }
            sb.AppendLine();

            sb.AppendLine("-----------------------------------------------------------------");
            sb.AppendLine($"命理大師：玉洞子 | 流年命書 v1.1 | {year} 年");
            return sb.ToString();
        }

        // ── analyze-topic ────────────────────────────────────────────────
        [HttpGet("analyze-topic")]
        [Authorize]
        public async Task<IActionResult> GetTopicAnalysis([FromQuery] string topic = "")
        {
            var validTopics = new[] { "事業", "婚姻", "財運", "子女", "父母", "兄妹", "學業", "買房", "投資", "住宅風水", "合夥", "出國", "開店", "健康" };
            if (string.IsNullOrWhiteSpace(topic) || !Array.Exists(validTopics, t => t == topic))
                return BadRequest(new { error = "請選擇有效的問事主題" });

            var identity = User.FindFirstValue(ClaimTypes.Email)
                         ?? User.FindFirstValue(ClaimTypes.Name)
                         ?? User.FindFirst("unique_name")?.Value;
            if (string.IsNullOrEmpty(identity))
                return Unauthorized(new { error = "請重新登入" });

            var user = await _context.Users.FirstOrDefaultAsync(u => u.UserName == identity || u.Email == identity);
            if (user == null) return BadRequest(new { error = "找不到用戶" });

            var userChart = await _context.UserCharts.FirstOrDefaultAsync(c => c.UserId == user.Id);
            if (userChart == null || string.IsNullOrEmpty(userChart.ChartJson))
                return BadRequest(new { error = "no_chart" });

            var (topicOk, topicErr) = await CheckTopicConsultAccess(user.Id);
            if (!topicOk) return BadRequest(new { error = topicErr });

            try
            {
                // ── 使用與流年/大運命書相同的計算方法，確保結論一致 ──
                var root = JsonDocument.Parse(userChart.ChartJson).RootElement;
                if (!root.TryGetProperty("bazi", out var bazi) && !root.TryGetProperty("baziInfo", out bazi))
                    return BadRequest(new { error = "命盤資料格式錯誤" });

                var yearP  = LfGetPillar(bazi, "yearPillar");
                var monthP = LfGetPillar(bazi, "monthPillar");
                var dayP   = LfGetPillar(bazi, "dayPillar");
                var timeP  = LfGetPillar(bazi, "timePillar");

                string yStem = LfPillarStem(yearP);   string yBranch = LfPillarBranch(yearP);
                string mStem = LfPillarStem(monthP);  string mBranch = LfPillarBranch(monthP);
                string dStem = LfPillarStem(dayP);    string dBranch = LfPillarBranch(dayP);
                string hStem = LfPillarStem(timeP);   string hBranch = LfPillarBranch(timeP);

                int birthYear = user.BirthYear ?? (DateTime.Today.Year - 30);
                string dmElem = KbStemToElement(dStem);
                var branches  = new[] { yBranch, mBranch, dBranch, hBranch };
                var wuXing    = LfCalcWuXingMatrix(yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch);
                double bodyPct   = LfGetBodyStrengthPct(dmElem, wuXing);
                string bodyLabel = LfGetBodyStrengthLabel(bodyPct);
                string season    = LfGetSeason(mBranch);
                var (pattern, yongShenElem, fuYiElem, _, tiaoHouElem5) = LfDetectGeJuAndYongShen(
                    yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch,
                    dmElem, wuXing, bodyPct, season);
                string jiShenElem = LfGetJiShenElem(yongShenElem, dmElem, bodyPct, pattern);
                var chartStems5   = new[] { yStem, mStem, dStem, hStem };

                var luckCycles = LfExtractLuckCycles(root);
                bool hasZiwei  = root.TryGetProperty("palaces", out var palaces) && palaces.ValueKind == JsonValueKind.Array;

                // 當前大運
                int currentYear = DateTime.Today.Year;
                int currentAge  = currentYear - birthYear;
                var curLuck = luckCycles.FirstOrDefault(lc => currentAge >= lc.startAge && currentAge < lc.endAge);
                string daiyunStem   = curLuck.stem   ?? "";
                string daiyunBranch = curLuck.branch ?? "";
                string daiyunSS     = !string.IsNullOrEmpty(daiyunStem) ? LfStemShiShen(daiyunStem, dStem) : "";

                // 近3年流年評分（與大運/流年命書完全相同演算法）
                var yearLines = new System.Text.StringBuilder();
                for (int y = currentYear; y <= currentYear + 2; y++)
                {
                    var (flStem, flBranch) = DyGetYearStemBranch(y);
                    int flAge = y - birthYear;
                    int flBaziScore  = DyCalcFlowYearBaziScore(flStem, flBranch, pattern, yongShenElem, fuYiElem, jiShenElem,
                        dmElem, bodyPct > 50, tiaoHouElem5, season, branches, chartStems5);
                    int flZiweiScore = hasZiwei ? DyCalcZiweiScore(flStem, palaces, daiyunStem, flAge) : 50;
                    string flCross   = DyCrossClass(flBaziScore, flZiweiScore);
                    string flSS      = LfStemShiShen(flStem, dStem);
                    yearLines.AppendLine($"  - {y}年 {flStem}{flBranch}（{flSS}）：{flCross}｜八字分={flBaziScore}，紫微分={flZiweiScore}，年齡={flAge}歲");
                }

                // 建立含 KB 預算數據的 prompt
                string kbFacts = $@"
### 本命盤命理系統預算數據（必須以此為唯一基礎，不可自行推翻）
- 格局：{pattern}
- 日主（{dStem}）五行：{dmElem}，身強弱：{bodyLabel}
- 用神元素：{yongShenElem}（對此命主有利的五行）
- 忌神元素：{jiShenElem}（對此命主有害的五行）
- 當前大運：{daiyunStem}{daiyunBranch}（十神：{daiyunSS}），{curLuck.startAge}-{curLuck.endAge}歲
- 近三年流年交叉評斷：
{yearLines}";

                string prompt = BuildTopicPrompt(userChart.ChartJson, topic, kbFacts);
                string aiResult = await CallGeminiApi(prompt);

                return Ok(new { result = aiResult });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "主題命書失敗 Topic={Topic} User={User}", topic, identity);
                return StatusCode(500, new { error = "主題命書生成失敗，請稍後再試", details = ex.Message });
            }
        }

        // ── Subscription discount helper ────────────────────────────────────
        // Returns discount rate (e.g. 0.8 for 20% off) for the given product type.
        // Returns 1.0 if no active subscription or no matching discount benefit.
        private async Task<(bool ok, string? error, int subscriptionId)> CheckSubscriptionQuota(string userId, string productCode)
        {
            var now = DateTime.UtcNow;
            var activeSub = await _context.UserSubscriptions
                .Include(s => s.Plan)
                    .ThenInclude(p => p.Benefits)
                .Where(s => s.UserId == userId && s.Status == "active" && s.ExpiryDate > now)
                .OrderByDescending(s => s.ExpiryDate)
                .FirstOrDefaultAsync();

            if (activeSub == null)
                return (false, "需要訂閱會員方案才能使用此功能", 0);

            var quotaBenefit = activeSub.Plan.Benefits
                .FirstOrDefault(b => b.ProductCode == productCode && b.BenefitType == "quota");

            if (quotaBenefit == null)
                return (false, "您的訂閱方案不包含此功能，請升級方案", 0);

            int quota = int.Parse(quotaBenefit.BenefitValue);
            int currentYear = DateTime.Now.Year;
            int usedCount = await _context.UserSubscriptionClaims
                .CountAsync(c => c.UserId == userId && c.ProductCode == productCode && c.ClaimYear == currentYear);

            if (usedCount >= quota)
                return (false, $"本年度命書已達使用次數上限（{quota} 次/年），請明年再使用", 0);

            return (true, null, activeSub.Id);
        }

        private async Task RecordSubscriptionClaim(string userId, int subscriptionId, string productCode)
        {
            _context.UserSubscriptionClaims.Add(new UserSubscriptionClaim
            {
                UserId = userId,
                SubscriptionId = subscriptionId,
                ProductCode = productCode,
                ClaimYear = DateTime.Now.Year,
                ClaimedAt = DateTime.UtcNow
            });
            await _context.SaveChangesAsync();
        }

        private async Task SaveUserReportAsync(string userId, string reportType, string title, string content, object? parameters = null)
        {
            try
            {
                string? paramsJson = parameters != null ? JsonSerializer.Serialize(parameters) : null;
                _context.UserReports.Add(new Ecanapi.Models.UserReport
                {
                    UserId = userId,
                    ReportType = reportType,
                    Title = title,
                    Content = content,
                    Parameters = paramsJson,
                    CreatedAt = DateTime.UtcNow,
                    Status = "pending_review"
                });
                await _context.SaveChangesAsync();

                // Notify admin of new report application
                var user = await _context.Users.FirstOrDefaultAsync(u => u.Id == userId);
                if (user != null)
                {
                    var frontendBase = _config["App:FrontendUrl"] ?? "https://yudongzi.tw";
                    var adminUrl = $"{frontendBase}/admin/reports";
                    _ = _email.SendAdminReportNotifyAsync(
                        user.UserName ?? userId,
                        user.Email ?? "",
                        title,
                        reportType,
                        adminUrl);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "SaveUserReport failed for user {UserId}", userId);
                // Non-critical: do not rethrow
            }
        }

        private async Task<(bool ok, string? error)> CheckTopicConsultAccess(string userId)
        {
            var now = DateTime.UtcNow;
            var activeSub = await _context.UserSubscriptions
                .Include(s => s.Plan)
                    .ThenInclude(p => p.Benefits)
                .Where(s => s.UserId == userId && s.Status == "active" && s.ExpiryDate > now)
                .OrderByDescending(s => s.ExpiryDate)
                .FirstOrDefaultAsync();

            if (activeSub == null)
                return (false, "需要訂閱會員方案才能使用此功能");

            bool hasAccess = activeSub.Plan.Benefits
                .Any(b => b.ProductCode == "TOPIC_CONSULT" && b.BenefitType == "access" && b.BenefitValue == "true");

            if (!hasAccess)
                return (false, "問事功能需要銀會員以上方案");

            return (true, null);
        }

        // ============================================================
        //  九星氣學加成區塊（純 KB，不走 Gemini）
        // ============================================================

        /// <summary>流年命書專用：命×運 + 命×流年 兩大核心組合</summary>
        private async Task<string> LnNsBuildSection(int birthYear, int birthMonth, int birthDay, int birthHour, int gender, int targetYear)
        {
            try
            {
                // 本命星
                int natalStar = Ecanapi.Services.NineStarCalcHelper.CalcNatalStar(birthYear, birthMonth, birthDay, gender);
                // 當前元運星（以流年所在運為準）
                int currentYun = Ecanapi.Services.NineStarCalcHelper.GetCurrentYun(targetYear);
                var (yunLabel, isProspering) = Ecanapi.Services.NineStarCalcHelper.GetStarYunStatus(natalStar, currentYun);
                // 通用流年星（不分男女，飛入中宮）
                int universalYearStar = Ecanapi.Services.NineStarCalcHelper.CalcYearStar(targetYear);
                // 飛宮：通用年星入中宮後，飛入本命宮位的星
                int flyingInNatal = Ecanapi.Services.NineStarCalcHelper.CalcFlyingStarInPalace(universalYearStar, natalStar);

                string natalName    = Ecanapi.Services.NineStarCalcHelper.StarNames[natalStar];
                string yunName      = Ecanapi.Services.NineStarCalcHelper.StarNames[currentYun];
                string uYearName    = Ecanapi.Services.NineStarCalcHelper.StarNames[universalYearStar];
                string flyingName   = Ecanapi.Services.NineStarCalcHelper.StarNames[flyingInNatal];

                // 查 KB 組合（先精確匹配，再 fallback 五行計算）
                var rules = await _context.NineStarCombinationRules.ToListAsync();

                (string title, string verdict, string description, string modified) LookupCombo(int starA, int starB)
                {
                    var named = rules.FirstOrDefault(r =>
                        r.StarA == starA && r.StarB == starB &&
                        (r.IsProspering == null || r.IsProspering == isProspering));

                    string t, v, d;
                    if (named != null)
                    {
                        t = named.Title;
                        v = named.Verdict;
                        d = named.Description;
                    }
                    else
                    {
                        var (rel, fv, note) = Ecanapi.Services.NineStarCalcHelper.CalcFiveElementCombination(starA, starB);
                        t = rel; v = fv; d = note;
                    }
                    string m = v == "特殊"
                        ? (isProspering ? "最吉（同星得運）" : "最凶（同星失運）")
                        : Ecanapi.Services.NineStarCalcHelper.ApplyYunModifier(v, isProspering);
                    return (t, v, d, m);
                }

                var (yunT, yunV, yunD, yunM) = LookupCombo(natalStar, currentYun);  // 命×運
                var (flyT, flyV, flyD, flyM) = LookupCombo(natalStar, flyingInNatal); // 命×流年（飛宮）

                var sb = new System.Text.StringBuilder();
                sb.AppendLine();
                sb.AppendLine("---");
                sb.AppendLine($"【第七章：九星氣學流年加成】");
                sb.AppendLine();
                sb.AppendLine($"本命星：{natalStar} {natalName}（{Ecanapi.Services.NineStarCalcHelper.StarAliases[natalStar]}）");
                sb.AppendLine($"當前元運：{currentYun}運（{yunName}），本命星為「{yunLabel}」");
                sb.AppendLine($"{targetYear} 年流年星：{universalYearStar} {uYearName}（飛入本命宮位者：{flyingInNatal} {flyingName}）");
                sb.AppendLine();

                // 命×運
                sb.AppendLine($"■ 命×運　{natalStar}{natalName} × {currentYun}{yunName}");
                if (!string.IsNullOrEmpty(yunT)) sb.AppendLine($"  組合名：{yunT}");
                sb.AppendLine($"  評斷：【{yunM}】");
                if (!string.IsNullOrEmpty(yunD)) sb.AppendLine($"  解析：{yunD}");
                sb.AppendLine();

                // 命×流年（飛宮盤：以飛入本命宮的星定吉凶）
                sb.AppendLine($"■ 命×流年（飛宮）　{natalStar}{natalName} × {flyingInNatal}{flyingName}");
                sb.AppendLine($"  飛宮原理：{targetYear} 年通用流年 {universalYearStar} {uYearName} 入中宮，逆飛後 {flyingInNatal} {flyingName} 飛入本命 {natalStar} 宮");
                if (!string.IsNullOrEmpty(flyT)) sb.AppendLine($"  組合名：{flyT}");
                sb.AppendLine($"  評斷：【{flyM}】");
                if (!string.IsNullOrEmpty(flyD)) sb.AppendLine($"  解析：{flyD}");
                sb.AppendLine();

                // 綜合結語
                string overallTone = (flyM.StartsWith("大吉") || flyM.StartsWith("最吉") || flyM.StartsWith("吉"))
                    ? "整體偏吉"
                    : (flyM.StartsWith("大凶") || flyM.StartsWith("最凶") || flyM.StartsWith("凶"))
                        ? "整體偏凶，需特別留意"
                        : "整體中平，宜穩健行事";

                string prosperNote = isProspering
                    ? $"本命星目前得運，{Ecanapi.Services.NineStarCalcHelper.StarProsper[natalStar]}"
                    : $"本命星目前失運，{Ecanapi.Services.NineStarCalcHelper.StarDecline[natalStar]}";

                sb.AppendLine($"九星流年綜評：命×運【{yunM}】，命×流年（飛宮）【{flyM}】，{overallTone}。");
                sb.AppendLine(prosperNote);
                sb.AppendLine();

                // 開運方位顏色
                string direction = Ecanapi.Services.NineStarCalcHelper.StarDirections[natalStar];
                string color     = Ecanapi.Services.NineStarCalcHelper.StarColors[natalStar];
                sb.AppendLine($"吉方位：{direction}　吉顏色：{color}");

                return sb.ToString();
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "LnNsBuildSection 生成失敗，略過九星章節");
                return string.Empty;
            }
        }

        /// <summary>根據生辰計算四柱九星並從 KB 生成加成區塊</summary>
        private async Task<string> NsBuildBirthSection(int birthYear, int birthMonth, int birthDay, int birthHour, int gender)
        {
            try
            {
                var ns = typeof(Ecanapi.Services.NineStarCalcHelper);

                int natalStar = Ecanapi.Services.NineStarCalcHelper.CalcNatalStar(birthYear, birthMonth, birthDay, gender);
                var birthDate = new DateTime(
                    birthYear,
                    Math.Clamp(birthMonth, 1, 12),
                    Math.Clamp(birthDay, 1, 28),
                    Math.Clamp(birthHour, 0, 23), 0, 0);
                int monthStar = Ecanapi.Services.NineStarCalcHelper.CalcMonthStar(birthDate);
                int dayStar   = Ecanapi.Services.NineStarCalcHelper.CalcDayStar(birthDate);
                int hourStar  = Ecanapi.Services.NineStarCalcHelper.CalcHourStar(birthDate);

                // 從 KB 查詢本命星特質
                var trait = await _context.NineStarTraits.FirstOrDefaultAsync(t => t.StarNumber == natalStar);

                string starAlias   = Ecanapi.Services.NineStarCalcHelper.StarAliases[natalStar];
                string starKeyword = Ecanapi.Services.NineStarCalcHelper.StarKeywords[natalStar];
                int currentYun     = Ecanapi.Services.NineStarCalcHelper.GetCurrentYun(DateTime.Today.Year);
                var (yunLabel, isProspering) = Ecanapi.Services.NineStarCalcHelper.GetStarYunStatus(natalStar, currentYun);

                // 東四命（1坎/3震/4巽/9離）vs 西四命（2坤/6乾/7兌/8艮）
                bool isEastGroup = natalStar == 1 || natalStar == 3 || natalStar == 4 || natalStar == 9;
                // 9運為離火（南方），依命星五行找相生方向作為住家座位建議
                string seatDirection;
                if (isEastGroup)
                {
                    // 東四命：坐東、坐北、坐南、坐東南為吉；9運南方旺火，離9命最佳
                    seatDirection = natalStar == 9 ? "坐南朝北（最佳，9運離火同氣相助）"
                        : natalStar == 3 || natalStar == 4 ? "坐東或坐東南（木生南方之火，與9運相生）"
                        : "坐北或坐東（水木互生，避開南方壓力）";
                }
                else
                {
                    // 西四命：坐西南、坐東北、坐西、坐西北為吉；9運火生土，2/8土命受生
                    seatDirection = natalStar == 2 || natalStar == 8 ? "坐東南或坐西南（9運火生土，受生最佳）"
                        : "坐西或坐西北（避開南方火剋，保持穩定）";
                }

                var sb = new System.Text.StringBuilder();
                sb.AppendLine();
                sb.AppendLine("---");
                sb.AppendLine("## 附錄：九星氣學加成");
                sb.AppendLine();

                sb.AppendLine($"### 宅命星解析：{Ecanapi.Services.NineStarCalcHelper.StarNames[natalStar]}（{starAlias}）");
                sb.AppendLine();
                sb.AppendLine(starKeyword);
                sb.AppendLine();
                sb.AppendLine($"**當前運勢狀態**：{currentYun}運（{Ecanapi.Services.NineStarCalcHelper.StarNames[currentYun]}），本命星目前為「{yunLabel}」");
                sb.AppendLine();

                if (isProspering)
                {
                    sb.AppendLine($"**得運展現**：{Ecanapi.Services.NineStarCalcHelper.StarProsper[natalStar]}");
                }
                else
                {
                    sb.AppendLine($"**失運注意**：{Ecanapi.Services.NineStarCalcHelper.StarDecline[natalStar]}");
                }
                sb.AppendLine();

                string direction = trait?.LuckyDirection ?? Ecanapi.Services.NineStarCalcHelper.StarDirections[natalStar];
                string color     = trait?.LuckyColor     ?? Ecanapi.Services.NineStarCalcHelper.StarColors[natalStar];
                int luckyNum     = trait?.LuckyNumber    ?? natalStar;

                sb.AppendLine("### 開運指引");
                sb.AppendLine();
                sb.AppendLine($"- **吉方位**：{direction}");
                sb.AppendLine($"  　面向或前往此方位，有助於提升個人氣場與運勢。");
                sb.AppendLine($"- **住家及辦公座位建議**：{seatDirection}");
                sb.AppendLine($"  　依宅命星所屬{(isEastGroup ? "東四命" : "西四命")}配合三元9運建議，此方向有助於提升運勢能量。");
                sb.AppendLine($"- **吉顏色**：{color}");
                sb.AppendLine($"  　穿著、配件或居家擺設選用此色系，可加持好運能量。");
                sb.AppendLine($"- **幸運數字**：{luckyNum}");
                sb.AppendLine($"  　選擇日期、樓層、電話號碼或車牌時，含有 {luckyNum} 的組合對您較有利。");

                if (natalStar == 5)
                    sb.AppendLine("- **特別提醒**：五黃土星能量強大，建議以白色、黃銅開運物件放置中宮位置，協助平衡能量。");

                return sb.ToString();
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "NsBuildBirthSection 生成失敗，略過九星區塊");
                return string.Empty;
            }
        }
    }

    public class AiRequest
    {
        public string Type { get; set; }
        public AstrologyRequest ChartRequest { get; set; }
        public int? TargetYear { get; set; }        // 流年命書用
        public string? Topic { get; set; }           // 問事用
        public int? FortuneDuration { get; set; }    // 大運命書年數：5/10/20/30/0=終身
    }

    public class CreateCheckoutSessionRequest
    {
        public string UserId { get; set; }
        public int Amount { get; set; }
        public int Points { get; set; }
        public string SuccessUrl { get; set; }
        public string CancelUrl { get; set; }
    }

    public class PointRecord
    {
        public int Id { get; set; }
        public string UserId { get; set; }
        public int Amount { get; set; }
        public string? Description { get; set; }
        public string? StripeSessionId { get; set; }
        public DateTime CreatedAt { get; set; }
    }
}