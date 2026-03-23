using Ecanapi.Data;
using Ecanapi.Models;
using Ecanapi.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Stripe;
using Stripe.Checkout;
using System.Security.Claims;
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
        private readonly IConfiguration _config;
        private readonly ILogger<ConsultationController> _logger;
        private static readonly HttpClient _httpClient = new HttpClient { Timeout = TimeSpan.FromMinutes(5) };

        public ConsultationController(IAstrologyService astrologyService, ApplicationDbContext context, IConfiguration config, ILogger<ConsultationController> logger)
        {
            _astrologyService = astrologyService;
            _context = context;
            _config = config;
            _logger = logger;
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
                return Ok(new { remainingPoints = user.Points });

            // 依類型與大運年數決定點數費用
            int cost = request.Type switch
            {
                "綜合性命書" or "綜合鑑定" => 50,
                "大運命書" => (request.FortuneDuration ?? 5) switch
                {
                    5 => 150,
                    10 => 200,
                    20 => 250,
                    30 => 300,
                    _ => 500   // 0 = 終身
                },
                "流年命書" => 20,
                "問事" => 20,
                _ => 50
            };

            string analyzeProductType = request.Type is "問事" ? "consultation" : "book";
            decimal analyzeDiscountRate = await GetSubscriptionDiscountRate(user.Id, analyzeProductType);
            int effectiveCost = (int)Math.Ceiling(cost * analyzeDiscountRate);

            if (user.Points < effectiveCost) return BadRequest(new { error = $"點數不足，此功能需要 {effectiveCost} 點" });

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

                user.Points -= effectiveCost;
                await _context.SaveChangesAsync();

                return Ok(new { result = aiResult, remainingPoints = user.Points });
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

            const int cost = 50;
            decimal kbDiscountRate = await GetSubscriptionDiscountRate(user.Id, "book");
            int kbEffectiveCost = (int)Math.Ceiling(cost * kbDiscountRate);
            if (user.Points < kbEffectiveCost)
                return BadRequest(new { error = $"點數不足，需要 {kbEffectiveCost} 點" });

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

                // 宮位四化（各宮宮干→化星飛入目標宮）
                var (mingLuPalace,  mingLuContent)  = await KbGongWeiSiHuaQuery(palaces, "命宮",  "化祿");
                var (mingJiPalace,  mingJiContent)  = await KbGongWeiSiHuaQuery(palaces, "命宮",  "化忌");
                var (offLuPalace,   offLuContent)   = await KbGongWeiSiHuaQuery(palaces, "官祿宮","化祿");
                var (offJiPalace,   offJiContent)   = await KbGongWeiSiHuaQuery(palaces, "官祿宮","化忌");
                var (wltLuPalace,   wltLuContent)   = await KbGongWeiSiHuaQuery(palaces, "財帛宮","化祿");
                var (wltJiPalace,   wltJiContent)   = await KbGongWeiSiHuaQuery(palaces, "財帛宮","化忌");
                var (spsLuPalace,   spsLuContent)   = await KbGongWeiSiHuaQuery(palaces, "夫妻宮","化祿");
                var (spsJiPalace,   spsJiContent)   = await KbGongWeiSiHuaQuery(palaces, "夫妻宮","化忌");
                var (hltJiPalace,   hltJiContent)   = await KbGongWeiSiHuaQuery(palaces, "疾厄宮","化忌");

                // 主星入宮（6/7/8三個主星文件）
                string starDescMing = await KbQueryStarInPalace(palaces, "命宮");
                string starDescOff  = await KbQueryStarInPalace(palaces, "官祿宮");
                string starDescWlt  = await KbQueryStarInPalace(palaces, "財帛宮");
                string starDescSps  = await KbQueryStarInPalace(palaces, "夫妻宮");
                string starDescHlt  = await KbQueryStarInPalace(palaces, "疾厄宮");

                // === 組裝命書 ===
                var sb_out = new StringBuilder();

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
                if (!string.IsNullOrEmpty(nianSiHuaXing)) sb_out.AppendLine($"【{nianGan}年干四化性格】{nianSiHuaXing}");
                if (!string.IsNullOrEmpty(ziweiMing))  sb_out.AppendLine($"【紫微格局綱領·{mingGongStars}】{ziweiMing}");
                if (!string.IsNullOrEmpty(starDescMing)) sb_out.AppendLine($"【命宮星情】{starDescMing}");
                // 先天四化與命格關聯
                if (!string.IsNullOrEmpty(siHuaLu))    sb_out.AppendLine($"【先天化祿·{siHuaLuPalace}】{siHuaLu}");
                if (!string.IsNullOrEmpty(siHuaQuan))  sb_out.AppendLine($"【先天化權·{siHuaQuanPalace}】{siHuaQuan}");
                if (!string.IsNullOrEmpty(siHuaJi))    sb_out.AppendLine($"【先天化忌·{siHuaJiPalace}】{siHuaJi}");
                // 命宮宮位四化飛出
                if (!string.IsNullOrEmpty(mingLuContent)) sb_out.AppendLine($"【命宮化祿飛{mingLuPalace}】{mingLuContent}");
                if (!string.IsNullOrEmpty(mingJiContent)) sb_out.AppendLine($"【命宮化忌飛{mingJiPalace}】{mingJiContent}");
                bool ch2BaziHas  = !string.IsNullOrEmpty(phenomenon) || !string.IsNullOrEmpty(rootType);
                bool ch2ZiweiHas = !string.IsNullOrEmpty(ziweiMing);
                if (ch2BaziHas && ch2ZiweiHas) sb_out.AppendLine("【格局交叉驗證】八字與紫微雙重印證，命格論斷可信度高。");
                sb_out.AppendLine();

                // --- 三、性格特質 ---
                sb_out.AppendLine("=== 三、性格特質 ===");
                sb_out.AppendLine("--- 八字性格論 ---");
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
                if (!string.IsNullOrEmpty(offLuContent)) sb_out.AppendLine($"【官祿宮化祿飛{offLuPalace}】{offLuContent}");
                if (!string.IsNullOrEmpty(offJiContent)) sb_out.AppendLine($"【官祿宮化忌飛{offJiPalace}】{offJiContent}");
                // 財帛宮12宮星情 + 宮位四化飛出 + 主星入宮
                if (!string.IsNullOrEmpty(ziweiWlt))     sb_out.AppendLine($"【財帛宮·{ziweiWltStar}】{ziweiWlt}");
                if (!string.IsNullOrEmpty(starDescWlt))  sb_out.AppendLine($"【財帛星性】{starDescWlt}");
                if (!string.IsNullOrEmpty(wltLuContent)) sb_out.AppendLine($"【財帛宮化祿飛{wltLuPalace}】{wltLuContent}");
                if (!string.IsNullOrEmpty(wltJiContent)) sb_out.AppendLine($"【財帛宮化忌飛{wltJiPalace}】{wltJiContent}");
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
                sb_out.AppendLine("=== 五、婚姻感情 ===");
                sb_out.AppendLine("--- 八字婚姻感情論 ---");
                if (!string.IsNullOrEmpty(aqfx))   sb_out.AppendLine($"【感情特質】{KbStripHtml(aqfx)}");
                if (!string.IsNullOrEmpty(astroX)) sb_out.AppendLine($"【婚姻論斷】{KbStripHtml(astroX)}");
                sb_out.AppendLine("--- 紫微婚姻感情論 ---");
                // 夫妻宮12宮星情 + 宮位四化飛出 + 主星入宮
                if (!string.IsNullOrEmpty(ziweiSps))     sb_out.AppendLine($"【夫妻宮·{ziweiSpsStar}】{ziweiSps}");
                if (!string.IsNullOrEmpty(starDescSps))  sb_out.AppendLine($"【夫妻星性】{starDescSps}");
                if (!string.IsNullOrEmpty(spsLuContent)) sb_out.AppendLine($"【夫妻宮化祿飛{spsLuPalace}】{spsLuContent}");
                if (!string.IsNullOrEmpty(spsJiContent)) sb_out.AppendLine($"【夫妻宮化忌飛{spsJiPalace}】{spsJiContent}");
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

                // --- 六、健康壽元 ---
                sb_out.AppendLine("=== 六、健康壽元 ===");
                sb_out.AppendLine("--- 八字健康壽元論 ---");
                if (!string.IsNullOrEmpty(jkfx))   sb_out.AppendLine($"【健康傾向】{KbStripHtml(jkfx)}");
                sb_out.AppendLine("--- 紫微健康壽元論 ---");
                // 疾厄宮12宮星情 + 化忌飛出 + 主星入宮
                if (!string.IsNullOrEmpty(ziweiHlt))     sb_out.AppendLine($"【疾厄宮·{ziweiHltStar}】{ziweiHlt}");
                if (!string.IsNullOrEmpty(starDescHlt))  sb_out.AppendLine($"【疾厄星性】{starDescHlt}");
                if (!string.IsNullOrEmpty(hltJiContent)) sb_out.AppendLine($"【疾厄宮化忌飛{hltJiPalace}】{hltJiContent}");
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
                if (!string.IsNullOrEmpty(astroZ)) sb_out.AppendLine($"【子息緣份】{KbStripHtml(astroZ)}");
                sb_out.AppendLine("--- 紫微家庭緣份論 ---");
                if (!string.IsNullOrEmpty(ziweiPar)) sb_out.AppendLine($"【父母宮·{ziweiParStar}】{ziweiPar}");
                if (!string.IsNullOrEmpty(ziweiCld)) sb_out.AppendLine($"【子女宮·{ziweiCldStar}】{ziweiCld}");
                if (!string.IsNullOrEmpty(siHuaJi) && (siHuaJiPalace == "父母宮" || siHuaJiPalace == "子女宮"))
                    sb_out.AppendLine($"【先天化忌警示·{siHuaJiPalace}】{siHuaJi}");
                sb_out.AppendLine();

                // 第八章大運流年已移除（綜合命書不含大運，另由大運命書提供）
                sb_out.AppendLine("-----------------------------------------------------------------");
                sb_out.AppendLine("命理鑑定大師：玉洞子  |  修身齊家，命在人心。  v3.0");

                user.Points -= kbEffectiveCost;
                await _context.SaveChangesAsync();

                return Ok(new { result = sb_out.ToString(), remainingPoints = user.Points });
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
            return string.Join("\n", lines.Skip(startIdx).Take(endIdx - startIdx)).Trim();
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

        // --- KB Helper: 過濾 ziwei 宮位內容，只保留命盤實際有的星才顯示的段落 ---
        // palaceStars: 該宮位自己的星（用於「同宮」類型檢查）
        // allChartStars: 整個命盤的星（用於「會照」「相夾」等跨宮檢查）
        private static string KbFilterZiweiContent(string content, HashSet<string> palaceStars, HashSet<string> allChartStars)
        {
            if (string.IsNullOrEmpty(content)) return content;
            var lines = content.Split('\n');
            var result = new StringBuilder();
            bool inConditional = false;
            bool includeSection = true;

            foreach (var rawLine in lines)
            {
                var line = rawLine.TrimEnd();
                var trigger = KbDetectSectionTriggerTyped(line);
                if (trigger != null)
                {
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
                }
                if (!inConditional || includeSection)
                    result.AppendLine(line);
            }
            return result.ToString().TrimEnd();
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

        // --- KB Helper: 依星曜縮寫找宮位名（返回正規化 XXX宮 格式）---
        private static string KbFindPalaceByStarAbbr(JsonElement palaces, string starAbbr)
        {
            if (palaces.ValueKind != JsonValueKind.Array || string.IsNullOrEmpty(starAbbr)) return "";
            foreach (var p in palaces.EnumerateArray())
            {
                string pname = p.TryGetProperty("palaceName", out var pn) ? pn.GetString() ?? "" :
                               p.TryGetProperty("name", out var n2) ? n2.GetString() ?? "" : "";
                foreach (var key in new[] { "majorStars", "mainStars", "secondaryStars" })
                {
                    if (p.TryGetProperty(key, out var stars) && stars.ValueKind == JsonValueKind.Array &&
                        stars.EnumerateArray().Any(s => (s.GetString() ?? "") == starAbbr))
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
            // 宮位四化.docx 格式：Title='宮位四化'，ConditionText 為空，
            // ResultText 以 "{源宮}{化type}入{目標宮短名}" 開頭（目標宮不一定帶"宮"字）
            string targetShort = targetPalace.TrimEnd('宮');
            string resultPrefix = $"{sourcePalaceName}{siHuaType}入{targetShort}";
            string content = await KbQuery($"SELECT COALESCE(\"ResultText\",'') AS \"Value\" FROM \"FortuneRules\" WHERE \"SourceFile\"='宮位四化.docx' AND \"Title\"='宮位四化' AND \"ResultText\" LIKE '{resultPrefix}%' LIMIT 1");
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

        private async Task<string> KbQueryStarInPalace(JsonElement palaces, string palaceName)
        {
            if (palaces.ValueKind != JsonValueKind.Array) return "";
            // 找到該宮的主星列表
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
                if (!StarInPalaceMap.TryGetValue(star, out var info)) continue;
                string result;
                if (info.useTitle)
                    result = await KbQuery($"SELECT COALESCE(\"ResultText\",'') AS \"Value\" FROM \"FortuneRules\" WHERE \"SourceFile\"='{info.file}' AND \"Title\" LIKE '{info.fullName}星入{palaceShort}%' LIMIT 1");
                else
                    result = await KbQuery($"SELECT COALESCE(\"ResultText\",'') AS \"Value\" FROM \"FortuneRules\" WHERE \"SourceFile\"='{info.file}' AND \"ResultText\" LIKE '{info.fullName}星入{palaceShort}宮：%' LIMIT 1");
                if (string.IsNullOrWhiteSpace(result)) continue;
                // ResultText style 含前綴 "七殺星入命宮：" → 去掉前綴取描述
                if (!info.useTitle && result.Contains('：'))
                    result = result.Substring(result.IndexOf('：') + 1);
                parts.AppendLine(result.Trim());
            }
            return parts.ToString().Trim();
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

            const int cost = 50;
            decimal lfDiscountRate = await GetSubscriptionDiscountRate(user.Id, "book");
            int lfEffectiveCost = (int)Math.Ceiling(cost * lfDiscountRate);
            if (user.Points < lfEffectiveCost)
                return BadRequest(new { error = $"點數不足，需要 {lfEffectiveCost} 點" });

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

                user.Points -= lfEffectiveCost;
                await _context.SaveChangesAsync();
                return Ok(new { result = report, luckCycles = cycleData, baziTable, yongJiTable, remainingPoints = user.Points });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "終身命書失敗 User={User}", identity);
                return StatusCode(500, new { error = "終身命書生成失敗，請稍後再試", details = ex.Message });
            }
        }

        [HttpGet("analyze-yudongzi")]
        [Authorize]
        public async Task<IActionResult> GetYudongziAnalysis()
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

                var chartStems2 = new[] { yStem, mStem, dStem, hStem };
                var scored = luckCycles.Select(lc =>
                {
                    int sc = LfCalcLuckScore(lc.stem, lc.branch, pattern, yongShenElem, fuYiElem, jiShenElem,
                        dmElem, bodyPct > 50, tiaoHouElem, season, branches, chartStems2, dStem);
                    return (lc.stem, lc.branch, lc.liuShen, lc.startAge, lc.endAge, score: sc, level: LfLuckLevel(sc));
                }).ToList();

                string report = LfBuildYudongziReport(
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

                // 管理員免費，不扣點
                return Ok(new { result = report, luckCycles = cycleData, baziTable, yongJiTable, remainingPoints = user.Points });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "玉洞子命書失敗 User={User}", identity);
                return StatusCode(500, new { error = "玉洞子命書生成失敗，請稍後再試", details = ex.Message });
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

            int cost = years switch { 5 => 150, 10 => 200, 20 => 250, 30 => 300, _ => 500 };
            decimal dyDiscountRate = await GetSubscriptionDiscountRate(user.Id, "book");
            int dyEffectiveCost = (int)Math.Ceiling(cost * dyDiscountRate);
            if (user.Points < dyEffectiveCost)
                return BadRequest(new { error = $"點數不足，需要 {dyEffectiveCost} 點" });

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
                        summary = DyYearSummary(crossClass, flStemSS, flBranchSS, baziScore, ziweiScore)
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
                if (hasZiwei)
                {
                    string ziweiPos = KbGetZiweiPosition(palaces);
                    ziweiFullContent = await KbZiweiFullQuery(palaces, ziweiPos);
                    chartStars = KbGetAllChartStars(palaces);
                }

                string report = DyBuildReport(
                    yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch,
                    yStemSS, mStemSS, hStemSS, yBranchSS, mBranchSS, dBranchSS, hBranchSS,
                    dmElem, wuXing, bodyPct, bodyLabel, season, seaLabel,
                    pattern, yongShenElem, fuYiElem, yongReason, jiShenElem,
                    luckCycles, annualDetails, hasZiwei, palaces, siHuaDescMap,
                    ziweiFullContent, chartStars,
                    gender, birthYear, years, branches, dStem);

                user.Points -= dyEffectiveCost;
                await _context.SaveChangesAsync();
                return Ok(new { result = report, annualForecasts, baziTable, luckCycles = scoredCycles, remainingPoints = user.Points });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "大運命書失敗 User={User}", identity);
                return StatusCode(500, new { error = "大運命書生成失敗，請稍後再試", details = ex.Message });
            }
        }

        // === Lf Static Data Tables ===

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
            ("木","春") => 1.8, ("木","夏") => 1.0, ("木","秋") => 0.2, ("木","冬") => 1.3, ("木","四季") => 0.5,
            ("火","春") => 1.3, ("火","夏") => 1.8, ("火","秋") => 0.5, ("火","冬") => 0.2, ("火","四季") => 1.0,
            ("土","春") => 0.2, ("土","夏") => 1.3, ("土","秋") => 1.0, ("土","冬") => 0.5, ("土","四季") => 1.8,
            ("金","春") => 0.5, ("金","夏") => 0.2, ("金","秋") => 1.8, ("金","冬") => 1.0, ("金","四季") => 1.3,
            ("水","春") => 1.0, ("水","夏") => 0.5, ("水","秋") => 1.3, ("水","冬") => 1.8, ("水","四季") => 0.2,
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
                if (!string.IsNullOrEmpty(elem)) scores[elem] += pts * LfSeasonMult(elem, season);
            }

            void AddBranch(string branch, double totalPts)
            {
                if (!LfBranchHiddenRatio.TryGetValue(branch, out var hidden)) return;
                double brMult = LfGetBranchMult(branch, branches);
                foreach (var (stem, ratio) in hidden)
                {
                    string elem = KbStemToElement(stem);
                    if (!string.IsNullOrEmpty(elem))
                        scores[elem] += totalPts * ratio * LfSeasonMult(elem, season) * brMult;
                }
            }

            AddStem(yStem, 10); AddStem(mStem, 10); AddStem(dStem, 10); AddStem(hStem, 10);
            AddBranch(yBranch, 10); AddBranch(dBranch, 10); AddBranch(hBranch, 10);
            AddBranch(mBranch, 30);

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
            >= 70 => "身強（極強）", >= 60 => "身強", >= 45 => "中和", >= 35 => "身弱", _ => "身弱（極弱）"
        };

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
            // 五行從旺格/從旺格：忌神=克日干之元素（破格之神）
            if (pattern == "從旺格" || LfWuXingGeJuSet.Contains(pattern))
                return LfElemOvercomeBy.GetValueOrDefault(dmElem, "");
            // 從強格：忌神=印/比劫（不可幫身對抗旺勢）
            if (pattern == "從強格")
                return dmElem;  // 比劫（自身力量，逆旺勢）
            // 從殺/從財/從兒格：忌神=印（生日主使其有力量對抗旺勢，破格之神）
            if (pattern is "從殺格" or "從財格" or "從兒格")
                return LfGenByElem.GetValueOrDefault(dmElem, "");  // 印星（最大破格威脅）
            // 身弱：大忌 = 克我（官殺），直接傷身
            // 身強：大忌 = 印星（生我讓身更旺，助力太多反被騙）
            if (bodyPct <= 40)
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

        private static readonly HashSet<string> LfWuXingGeJuSet =
            new() { "曲直格", "炎上格", "稼穡格", "從革格", "潤下格" };

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

            string chosenStem = "";
            if (LfBranchHiddenRatio.TryGetValue(mBranch, out var mH) && mH.Count > 0)
            {
                // 找出所有「透干」的月支藏干
                var transparentHidden = mH.Where(h => allHeavenStems.Contains(h.stem)).ToList();

                if (transparentHidden.Count == 0)
                {
                    // 無透干：取月支藏干中根氣最強者
                    chosenStem = mH.OrderByDescending(h => LfStemRootScore(h.stem, allBranches)).First().stem;
                }
                else if (transparentHidden.Count == 1)
                {
                    // 單一透干：直接取
                    chosenStem = transparentHidden[0].stem;
                }
                else
                {
                    // 多干透出：調候優先，其次根氣
                    string? tiaoHouMatch = tiaoHouList.FirstOrDefault(t => transparentHidden.Any(h => h.stem == t));
                    chosenStem = tiaoHouMatch
                        ?? transparentHidden.OrderByDescending(h => LfStemRootScore(h.stem, allBranches)).First().stem;
                }
            }
            string chosenSS = LfStemShiShen(chosenStem, dStem);

            // Rule 4: 比劫不取八格，改外格（建祿格/月刃格）
            string pattern = chosenSS switch
            {
                "正官" => "正官格", "七殺" => "七殺格", "正印" => "正印格", "偏印" => "偏印格",
                "正財" => "正財格", "偏財" => "偏財格", "食神" => "食神格", "傷官" => "傷官格",
                "比肩" => "建祿格", "劫財" => "月刃格", _ => "普通格"
            };

            // Check 五行從旺外格（優先於一般外格判斷）
            string wuXingGeJu = LfDetectWuXingGeJu(
                dmElem, mBranch, allHeavenStems, new[] { yBranch, mBranch, dBranch, hBranch });
            if (!string.IsNullOrEmpty(wuXingGeJu))
                pattern = wuXingGeJu;

            // Check 從格/從旺格（體極弱/極強時順旺勢）
            if (bodyPct <= 20)
            {
                string guanElemD = LfElemOvercomeBy.GetValueOrDefault(dmElem, ""); // 七殺（克日主）
                string caiElemD  = LfElemOvercome.GetValueOrDefault(dmElem, "");   // 財星（日主克）
                string shiElemD  = LfElemGen.GetValueOrDefault(dmElem, "");        // 食傷（日主生）
                double guanPctD  = wuXing.GetValueOrDefault(guanElemD, 0);
                double caiPctD   = wuXing.GetValueOrDefault(caiElemD, 0);
                double shiPctD   = wuXing.GetValueOrDefault(shiElemD, 0);
                double oppPct    = guanPctD + caiPctD + shiPctD;
                // 日主極弱，從格判定：按最強元素決定從格種類（非一律從強格）
                if (oppPct >= 70)
                {
                    if      (guanPctD >= caiPctD && guanPctD >= shiPctD) pattern = "從殺格";
                    else if (caiPctD  >= guanPctD && caiPctD >= shiPctD) pattern = "從財格";
                    else                                                  pattern = "從兒格";
                }
            }
            else if (bodyPct >= 80)
            {
                double sameElem = wuXing.GetValueOrDefault(dmElem, 0) + wuXing.GetValueOrDefault(LfGenByElem.GetValueOrDefault(dmElem, ""), 0);
                if (sameElem >= 75) pattern = "從旺格";
            }

            string yongShenElem;
            string fuYiElem;  // secondary: 扶抑用神，身弱+調候時補充比劫/印
            string reason;

            // 扶抑用神（不受調候影響）
            string fuYiElemCalc;
            if (pattern == "從強格")
                fuYiElemCalc = new[] { LfElemGen.GetValueOrDefault(dmElem,""), LfElemOvercome.GetValueOrDefault(dmElem,""), LfElemOvercomeBy.GetValueOrDefault(dmElem,"") }
                    .OrderByDescending(e => wuXing.GetValueOrDefault(e, 0)).First();
            else if (pattern == "從旺格" || LfWuXingGeJuSet.Contains(pattern))
                fuYiElemCalc = dmElem;  // 五行從旺/從旺格：用神=日干本元素
            else if (bodyPct >= 60)
            {
                string outElem  = LfElemGen.GetValueOrDefault(dmElem, "");        // 食傷（洩秀）
                string caiElem  = LfElemOvercome.GetValueOrDefault(dmElem, "");   // 財星
                string guanElem = LfElemOvercomeBy.GetValueOrDefault(dmElem, ""); // 官殺
                // 月令格局優先：印格/食傷格 → 食傷洩秀；比劫格 → 官殺或食傷
                if (pattern is "偏印格" or "正印格" or "食神格" or "傷官格")
                    fuYiElemCalc = outElem;
                else if (pattern is "建祿格" or "月刃格")
                    fuYiElemCalc = (LfElemHasRoot(guanElem, yBranch, mBranch, dBranch, hBranch)
                                  || wuXing.GetValueOrDefault(guanElem, 0) >= 10) ? guanElem : outElem;
                else
                    fuYiElemCalc = wuXing.GetValueOrDefault(guanElem, 0) >= 10 ? guanElem : caiElem;
                // 有根優先：若月令所選無根但食傷有根，改用食傷
                if (fuYiElemCalc != outElem
                    && !LfElemHasRoot(fuYiElemCalc, yBranch, mBranch, dBranch, hBranch)
                    && wuXing.GetValueOrDefault(fuYiElemCalc, 0) < 10
                    && LfElemHasRoot(outElem, yBranch, mBranch, dBranch, hBranch))
                    fuYiElemCalc = outElem;
            }
            else
            {
                string inElem = LfGenByElem.GetValueOrDefault(dmElem, "");
                fuYiElemCalc = wuXing.GetValueOrDefault(inElem, 0) >= 10 ? inElem : dmElem;
            }

            // 扶抑為主
            yongShenElem = fuYiElemCalc;
            if (bodyPct >= 60)
            {
                string outElemR = LfElemGen.GetValueOrDefault(dmElem, "");
                reason = yongShenElem == outElemR
                    ? $"扶抑法（身強{pattern}，取食傷洩秀）"
                    : $"扶抑法（身強{pattern}，取官殺/財洩耗）";
            }
            else
            {
                reason = bodyPct <= 40 ? "扶抑法（身弱，取印比生扶）" : "中和格（月令用神為主）";
            }
            if (pattern == "從強格") reason = "從強格（順旺勢）";
            else if (pattern == "從旺格") reason = "從旺格（順旺勢）";
            else if (LfWuXingGeJuSet.Contains(pattern)) reason = $"{pattern}（五行純粹，順旺勢）";

            // fuYiElem = 另一個扶身元素（身弱：印/比劫互補；身強：官/財互補）
            string inElemLocal  = LfGenByElem.GetValueOrDefault(dmElem, "");
            string guanElemLocal = LfElemOvercomeBy.GetValueOrDefault(dmElem, "");
            string caiElemLocal  = LfElemOvercome.GetValueOrDefault(dmElem, "");
            if (pattern is "從強格" or "從旺格" || LfWuXingGeJuSet.Contains(pattern))
                fuYiElem = yongShenElem;
            else if (bodyPct <= 40)
                fuYiElem = yongShenElem == inElemLocal ? dmElem : inElemLocal;   // 印/比劫互補
            else if (bodyPct >= 60)
            {
                string outElemL = LfElemGen.GetValueOrDefault(dmElem, "");
                if (yongShenElem == outElemL)            fuYiElem = caiElemLocal;      // 食傷→fuYi=財
                else if (yongShenElem == caiElemLocal)   fuYiElem = outElemL;          // 財→fuYi=食傷
                else                                     fuYiElem = caiElemLocal;      // 官殺→fuYi=財
            }
            else
                fuYiElem = yongShenElem;

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
            // 附加調候補用神到 reason（若與主用神不同才標注）
            if (!string.IsNullOrEmpty(tiaoHouElem) && tiaoHouElem != yongShenElem)
                reason += $"；調候補用：{tiaoHouElem}";
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
            string wx = $"木{wuXing["木"]:F0}% 火{wuXing["火"]:F0}% 土{wuXing["土"]:F0}% 金{wuXing["金"]:F0}% 水{wuXing["水"]:F0}%";
            int currentAge = DateTime.Today.Year - birthYear;

            sb.AppendLine("=================================================================");
            sb.AppendLine("                         八 字 命 書");
            sb.AppendLine("=================================================================");
            sb.AppendLine();

            // === Ch.1 命盤基本資訊 ===
            sb.AppendLine("【第一章：命盤基本資訊】");
            sb.AppendLine($"性別：{genderText}  出生年：{birthYear} 年");
            sb.AppendLine($"四柱：{yStem}{yBranch} {mStem}{mBranch} {dStem}{dBranch} {hStem}{hBranch}");
            sb.AppendLine($"十神：年干{SS(yStemSS)} 年支{SS(yBranchSS)} 月干{SS(mStemSS)} 月支{SS(mBranchSS)} 時干{SS(hStemSS)} 時支{SS(hBranchSS)}");
            sb.AppendLine($"日主：{dStem}（{dmElem}）");
            if (scored.Count > 0)
                sb.AppendLine($"大運起運：{scored[0].startAge} 歲，共 {scored.Count} 步大運");
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
            sb.AppendLine($"比印陣：{biJiPct:F0}% | 洩克陣：{100 - biJiPct:F0}%");
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
            sb.AppendLine(LfBuildYongJiTable(yongShenElem, fuYiElem, jiShenElem, tuneElemDisp, dStem, branches));
            sb.AppendLine();

            // === Ch.5 六親論斷 ===
            sb.AppendLine("【第五章：六親論斷（量化版）】");
            int yrSc  = LfCalcRelativeScore(yStemSS, yBranchSS, yongShenElem, jiShenElem, yStem, yBranch, branches);
            int moSc  = LfCalcRelativeScore(mStemSS, mBranchSS, yongShenElem, jiShenElem, mStem, mBranch, branches);
            int daySc = LfCalcRelativeScore(dBranchSS, dBranchSS, yongShenElem, jiShenElem, dStem, dBranch, branches);
            int hrSc  = LfCalcRelativeScore(hStemSS, hBranchSS, yongShenElem, jiShenElem, hStem, hBranch, branches);
            sb.AppendLine($"年柱·出身祖業（{yStem}{yBranch}）：緣分 {yrSc} 分（{LfRelativeLevel(yrSc)}）");
            sb.AppendLine($"  {LfYearDesc(yrSc)}");
            if (!LfShouldSkipPalace("父母宮", currentAge))
            {
                sb.AppendLine($"月柱·父母兄弟（{mStem}{mBranch}）：緣分 {moSc} 分（{LfRelativeLevel(moSc)}）");
                sb.AppendLine($"  {LfMonthDesc(moSc)}");
            }
            else
            {
                sb.AppendLine($"月柱·{LfPalaceAgeLabel("父母宮", currentAge)}（{mStem}{mBranch}）：緣分 {moSc} 分（{LfRelativeLevel(moSc)}）");
                sb.AppendLine($"  {LfMonthDesc(moSc)}");
            }
            if (!LfShouldSkipPalace("夫妻宮", currentAge))
            {
                sb.AppendLine($"日支·配偶緣分（{dBranch}）：緣分 {daySc} 分（{LfRelativeLevel(daySc)}）");
                sb.AppendLine($"  {LfDayDesc(daySc, gender)}");
            }
            if (!LfShouldSkipPalace("子女宮", currentAge))
            {
                sb.AppendLine($"時柱·子女晚運（{hStem}{hBranch}）：緣分 {hrSc} 分（{LfRelativeLevel(hrSc)}）");
                sb.AppendLine($"  {LfHourDesc(hrSc)}");
            }
            else
            {
                sb.AppendLine($"時柱·{LfPalaceAgeLabel("子女宮", currentAge)}（{hStem}{hBranch}）：緣分 {hrSc} 分（{LfRelativeLevel(hrSc)}）");
                sb.AppendLine($"  {LfHourDesc(hrSc)}");
            }
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
            sb.AppendLine($"格局：{pattern}，天生適合{LfCareerDesc(pattern)}。");
            double caiPct = wuXing.GetValueOrDefault(LfElemOvercome.GetValueOrDefault(dmElem, ""), 0);
            sb.AppendLine($"財星（{LfElemOvercome.GetValueOrDefault(dmElem,"")}）占命局 {caiPct:F0}%，日主強弱 {bodyPct:F0}%。");
            sb.AppendLine(LfWealthDesc(caiPct, bodyPct, wuXing.GetValueOrDefault(dmElem, 0)));
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
                sb.AppendLine();
            }

            // === Ch.9 健康壽元 ===
            sb.AppendLine("【第九章：健康壽元】");
            sb.AppendLine($"命局五行：{wx}");
            sb.AppendLine(LfHealthDesc(wuXing, seaLabel));
            sb.AppendLine();

            // === Ch.10 大運逐運 ===
            sb.AppendLine("【第十章：大運逐運論斷（百分制評分）】");
            {
                string ageHint = LfAgeTopicHint(currentAge);
                if (!string.IsNullOrEmpty(ageHint)) sb.AppendLine(ageHint);
            }
            string[] branchSSArr = { yBranchSS, mBranchSS, dBranchSS, hBranchSS };
            foreach (var c in scored)
            {
                string lcSS = LfStemShiShen(c.stem, dStem);
                string lcBranchMs = LfBranchHiddenRatio.TryGetValue(c.branch, out var lcBH) && lcBH.Count > 0 ? lcBH[0].stem : "";
                string lcBranchSS = !string.IsNullOrEmpty(lcBranchMs) ? LfStemShiShen(lcBranchMs, dStem) : "";
                sb.AppendLine($"{c.startAge}-{c.endAge} 歲 大運：{c.stem}{c.branch}（天干{lcSS}·地支{lcBranchSS}）  評分：{c.score} 分（{c.level}）");
                sb.AppendLine($"  {LfLuckDesc(c.score, c.level)}");
                string events = LfBranchEventsPalace(c.branch, lcBranchSS, branches, branchSSArr, c.startAge);
                if (!string.IsNullOrEmpty(events))
                {
                    sb.AppendLine($"  【地支事項】大運地支{c.branch}（{lcBranchSS}）：");
                    sb.AppendLine($"  {events}");
                }
            }
            sb.AppendLine();

            // === Ch.11 流年重點 ===
            sb.AppendLine("【第十一章：流年重點吉凶】");
            sb.AppendLine(LfKeyYears(scored, birthYear, yongShenElem, jiShenElem));
            sb.AppendLine();

            // === Ch.12 總評 ===
            sb.AppendLine("【第十二章：一生命運總評】");
            double earlyAvg = scored.Where(c => c.endAge <= 30).Select(c => (double)c.score).DefaultIfEmpty(50).Average();
            double midAvg   = scored.Where(c => c.startAge >= 31 && c.endAge <= 50).Select(c => (double)c.score).DefaultIfEmpty(50).Average();
            double lateAvg  = scored.Where(c => c.startAge > 50).Select(c => (double)c.score).DefaultIfEmpty(50).Average();
            sb.AppendLine($"前運（0-30 歲）平均 {earlyAvg:F0} 分：{LfPeriodDesc(earlyAvg)}");
            sb.AppendLine($"中運（31-50 歲）平均 {midAvg:F0} 分：{LfPeriodDesc(midAvg)}");
            sb.AppendLine($"後運（51 歲後）平均 {lateAvg:F0} 分：{LfPeriodDesc(lateAvg)}");
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

            sb.AppendLine("-----------------------------------------------------------------");
            sb.AppendLine("命理大師：玉洞子 | 八字命書 v2.1");
            return sb.ToString();
        }

        // === 玉洞子命書（八章版，管理員內部）===

        private static string LfBuildYudongziReport(
            string yStem, string yBranch, string mStem, string mBranch,
            string dStem, string dBranch, string hStem, string hBranch,
            string yStemSS, string mStemSS, string hStemSS,
            string yBranchSS, string mBranchSS, string dBranchSS, string hBranchSS,
            string dmElem, Dictionary<string, double> wuXing, double bodyPct, string bodyLabel,
            string season, string seaLabel, string pattern, string yongShenElem, string fuYiElem,
            string yongReason, string jiShenElem,
            List<(string stem, string branch, string liuShen, int startAge, int endAge, int score, string level)> scored,
            int gender, int birthYear)
        {
            var sb = new StringBuilder();
            string genderText = gender == 1 ? "男（乾造）" : "女（坤造）";
            var branches = new[] { yBranch, mBranch, dBranch, hBranch };
            string SS(string ss) => string.IsNullOrEmpty(ss) ? "" : $"（{ss}）";
            string wx = $"木{wuXing["木"]:F0}% 火{wuXing["火"]:F0}% 土{wuXing["土"]:F0}% 金{wuXing["金"]:F0}% 水{wuXing["水"]:F0}%";
            int currentAgeYdz = DateTime.Today.Year - birthYear;

            sb.AppendLine("=================================================================");
            sb.AppendLine("                      玉 洞 子 命 書");
            sb.AppendLine("=================================================================");
            sb.AppendLine();

            // === Ch.1 命盤基本資訊 ===
            sb.AppendLine("【第一章：命盤基本資訊】");
            sb.AppendLine($"性別：{genderText}  出生年：{birthYear} 年");
            sb.AppendLine($"四柱：{yStem}{yBranch} {mStem}{mBranch} {dStem}{dBranch} {hStem}{hBranch}");
            sb.AppendLine($"十神：年干{SS(yStemSS)} 年支{SS(yBranchSS)} 月干{SS(mStemSS)} 月支{SS(mBranchSS)} 時干{SS(hStemSS)} 時支{SS(hBranchSS)}");
            sb.AppendLine($"日主：{dStem}（{dmElem}）");
            if (scored.Count > 0)
                sb.AppendLine($"大運起運：{scored[0].startAge} 歲，共 {scored.Count} 步大運");
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
            sb.AppendLine($"比印陣：{biJiPct:F0}% | 洩克陣：{100 - biJiPct:F0}%");
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
                sb.AppendLine($"次忌(△忌)：{jiYongElemDisp}，天干 {LfElemStems(jiYongElemDisp)}，地支 {LfElemBranches(jiYongElemDisp)}（克用神{yongShenElem}）");
            sb.AppendLine($"格局說明：{LfPatternDesc(pattern)}");
            sb.AppendLine();
            sb.AppendLine(LfBuildYongJiTable(yongShenElem, fuYiElem, jiShenElem, tuneElemDisp, dStem, branches));
            sb.AppendLine();

            // === Ch.5 事業格局鑑定 ===
            sb.AppendLine("【第五章：事業格局鑑定】");
            sb.AppendLine();
            sb.AppendLine(KbSanmenCareer(
                yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch,
                yStemSS, mStemSS, hStemSS, yBranchSS, mBranchSS, dBranchSS, hBranchSS,
                dmElem, pattern, bodyPct, yongShenElem, jiShenElem, wuXing));

            // === Ch.6 六親緣分·婚姻深度鑑定 ===
            sb.AppendLine("【第六章：六親緣分·婚姻深度鑑定】");
            sb.AppendLine();
            {
                string ageHintYdz = LfAgeTopicHint(currentAgeYdz);
                if (!string.IsNullOrEmpty(ageHintYdz)) sb.AppendLine(ageHintYdz);
            }
            sb.AppendLine(KbSanmenSixRelatives(
                yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch,
                yStemSS, mStemSS, hStemSS, yBranchSS, mBranchSS, dBranchSS, hBranchSS,
                dmElem, pattern, bodyPct, yongShenElem, jiShenElem, wuXing, gender, birthYear, scored));

            // === Ch.7 疾厄壽元鑑定 ===
            sb.AppendLine("【第七章：疾厄壽元鑑定】");
            sb.AppendLine();
            sb.AppendLine(KbSanmenHealthLongevity(
                yStem, mStem, hStem, yBranch, mBranch, dBranch, hBranch,
                dStem, dmElem, bodyPct, yongShenElem, jiShenElem,
                wuXing, season, seaLabel, scored));

            // === Ch.8 居家風水鑑定 ===
            sb.AppendLine("【第八章：居家風水鑑定】");
            sb.AppendLine();
            sb.AppendLine(KbSanmenFengShui(
                yStem, yBranch, mStem, mBranch, dStem, dBranch, hStem, hBranch,
                dmElem, bodyPct, yongShenElem, jiShenElem, wuXing, scored));

            sb.AppendLine("-----------------------------------------------------------------");
            sb.AppendLine("命理大師：玉洞子 | 玉洞子命書 v1.0（內部版）");
            return sb.ToString();
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

            // 五行職業取象
            sb.AppendLine();
            sb.AppendLine("【五行職業取象（適合行業）】");
            sb.Append(KbSanmenJobByElem(dmElem, pattern, yongShenElem, isHuangliang, isYangZhiYin));

            return sb.ToString();
        }

        private static bool IsGuanSha(string ss) => ss == "正官" || ss == "七殺";
        private static bool IsYin(string ss)     => ss == "正印" || ss == "偏印";

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
            sb.AppendLine("【婚期推算（大運時機）】");
            var marriageLucks = scored.Where(lc =>
            {
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
            List<(string stem, string branch, string liuShen, int startAge, int endAge, int score, string level)> scored)
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
            sb.AppendLine("【大運健康風險期】");
            var riskLucks = scored.Where(lc =>
            {
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

        private static string LfElemStems(string elem) => elem switch
        { "木"=>"甲乙","火"=>"丙丁","土"=>"戊己","金"=>"庚辛","水"=>"壬癸",_=>"" };

        private static string LfElemBranches(string elem) => elem switch
        { "木"=>"寅卯辰","火"=>"巳午未","土"=>"辰戌丑未","金"=>"申酉戌","水"=>"亥子丑",_=>"" };

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
            string baseChar = dmElem switch
            {
                "木" => "仁慈上進，有理想有原則，固執但有方向感。",
                "火" => "熱情開朗，重視禮儀，感情豐富，但性子急躁。",
                "土" => "穩重誠信，待人厚道，保守踏實，但反應較慢。",
                "金" => "果決義氣，剛強有原則，重然諾，但有時過於強硬。",
                "水" => "聰明靈活，多謀善慮，思維廣博，但有時多慮患得患失。",
                _ => ""
            };
            string strChar = bodyPct >= 60 ? "個性較強硬主觀，自信有領導慾，宜學習傾聽。" :
                             bodyPct <= 40 ? "個性較保守依賴，情緒易受外界影響，宜培養自信。" :
                             "個性均衡適應力強，能隨機應變，圓融處世。";
            string patChar = pattern switch
            {
                "正官格" => "規矩守法，重視名譽，適合在制度環境中發揮。",
                "七殺格" => "魄力十足，敢衝敢拚，但須注意控制衝動。",
                "食神格" => "喜享受，有藝術氣質，待人隨和。",
                "傷官格" => "聰明才俊，有叛逆性，不喜受約束。",
                "正財格" => "務實重物質，勤勞踏實，穩健理財。",
                _ => ""
            };
            return $"{baseChar} {strChar}{(string.IsNullOrEmpty(patChar) ? "" : " " + patChar)}";
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
            if (caiPct >= 25 && bodyPct >= 60) return "財多身強，財富豐厚，一生衣食無憂，能守住財富。";
            if (caiPct >= 25 && bodyPct < 40)  return "財多身弱，雖有財路但難以守住，辛苦奔波，錢財易散，宜節流。";
            if (biPct >= 30 && caiPct < 15)    return "比劫奪財，財路受阻，合夥易損，獨立經營較佳，防借財給人。";
            if (caiPct < 15 && bodyPct >= 60)  return "財星偏弱，守成有餘進財不足，宜穩健理財，不宜冒進。";
            return "財富中等，靠自身努力積累，運勢起伏時注意守財。";
        }

        private static string LfMarriageDesc(double spousePct, string dBranch, string dStem, string dmElem, int gender, string[] branches)
        {
            string star = gender == 1 ? "妻緣" : "夫緣";
            bool anyChong = branches.Where(b => b != dBranch).Any(b => LfChong.Contains(dBranch + b));
            string result = spousePct >= 20
                ? $"{star}不弱，感情豐富，異性緣佳，婚姻較有依靠。"
                : $"{star}偏薄，感情波折較多，緣分需珍惜，婚姻宜謹慎選擇。";
            if (anyChong) result += " 日支逢沖，婚姻有波折，夫妻易有摩擦或分離之象，需多包容。";
            string dBranchMainElem = LfBranchHiddenRatio.TryGetValue(dBranch, out var bhe) && bhe.Count > 0 ? KbStemToElement(bhe[0].stem) : "";
            string dStemElem = KbStemToElement(dStem);
            if (gender == 1 && LfElemOvercome.GetValueOrDefault(dStemElem, "") == dBranchMainElem)
                result += " 日干克日支，婚後夫妻個性有差異，需多體諒包容。";
            return result;
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
            >= 70 => "整體走勢佳，運勢蒸蒸日上，把握機遇積極發展。",
            >= 50 => "整體平穩，有起有伏，宜穩中求進。",
            _ => "整體考驗較多，宜低調蓄積，靜待轉機。"
        };

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
            string action = crossClass switch
            {
                "大吉" => "宜積極進取，把握機遇。",
                "吉"   => "整體向好，宜穩健前進。",
                "平"   => "平穩為主，守成待機。",
                "小凶" => "宜謹慎保守，避免冒進。",
                "大凶" => "宜低調守成，防範風險。",
                _      => "平穩行事。"
            };
            string ssDesc = !string.IsNullOrEmpty(flStemSS) ? $"（{flStemSS}年）" : "";
            return $"{baziDesc}、{ziweiDesc}，{crossClass}{ssDesc}。{action}";
        }

        private static string DyCrossDesc(string crossClass, string flStemSS, string flBranchSS, int baziScore, int ziweiScore)
        {
            string desc = crossClass switch
            {
                "大吉" => "八字用神得力、紫微吉曜臨宮，雙重加持，為黃金年份。宜積極進取，把握機遇，無論事業投資婚姻皆可有所作為。",
                "吉"   => "八字與紫微均向好，整體運勢順暢。宜穩健前進，趁勢佈局，擴展人脈與資源。",
                "平"   => "八字與紫微各有吉凶相抵，整體平穩。宜守成待機，不急進不冒險，積累實力。",
                "小凶" => "八字或紫微有一方出現凶象，需提高警覺。宜謹慎保守，避免重大決策，做好風險管理。",
                "大凶" => "八字忌神活躍、紫微凶曜臨宮，雙重壓力，需謹慎應對。宜低調守成，避免冒進，積極化解，以守為攻。",
                _      => "平穩行事，量力而為。"
            };
            string ssNote = "";
            if (!string.IsNullOrEmpty(flStemSS)) ssNote = $"流年天干為{flStemSS}";
            if (!string.IsNullOrEmpty(flBranchSS)) ssNote += (ssNote.Length > 0 ? "、" : "") + $"地支藏{flBranchSS}";
            return string.IsNullOrEmpty(ssNote) ? desc : $"（{ssNote}）{desc}";
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
            string[] dyTiaoHouList = LfTiaoHou.TryGetValue(dStemRef, out var dyTh1) && dyTh1.TryGetValue(mBranch, out var dyTh2)
                ? dyTh2 : Array.Empty<string>();
            string dyTiaoHouElem = dyTiaoHouList.Length > 0 ? KbStemToElement(dyTiaoHouList[0]) : "";
            var coveredLucks = luckCycles.Where(lc =>
                annualDetails.Any(a => a.age >= lc.startAge && a.age < lc.endAge)).ToList();
            if (coveredLucks.Count == 0) coveredLucks = luckCycles.Take(2).ToList();
            foreach (var lc in coveredLucks)
            {
                string lcSS  = LfStemShiShen(lc.stem, dStemRef);
                string lcBMs = LfBranchHiddenRatio.TryGetValue(lc.branch, out var lcBH2) && lcBH2.Count > 0 ? lcBH2[0].stem : "";
                string lcBSS = !string.IsNullOrEmpty(lcBMs) ? LfStemShiShen(lcBMs, dStemRef) : "";
                int lcScore  = LfCalcLuckScore(lc.stem, lc.branch, pattern, yongShenElem, fuYiElem, jiShenElem,
                    dmElem, bodyPct > 50, dyTiaoHouElem, season, branches, dyChartStems, dStemRef);
                sb.AppendLine($"{lc.startAge}-{lc.endAge} 歲 大運：{lc.stem}{lc.branch}（天干{lcSS}·地支{lcBSS}）  評分：{lcScore} 分（{LfLuckLevel(lcScore)}）");
                sb.AppendLine($"  {LfLuckDesc(lcScore, LfLuckLevel(lcScore))}");
                string palaceEvents = LfBranchEventsPalace(lc.branch, lcBSS, branches, branchSSArr, lc.startAge);
                if (!string.IsNullOrEmpty(palaceEvents))
                {
                    sb.AppendLine($"  【地支事項】大運地支{lc.branch}（{lcBSS}）：");
                    sb.AppendLine($"  {palaceEvents}");
                }
                // 大限宮干化忌入關鍵宮位警示（早於詳細宮位分析呈現，讓讀者優先注意重大風險）
                if (hasZiwei)
                {
                    var warnPalaces = DyGetOverlappingDecadePalaces(palaces, lc.startAge, lc.endAge);
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
                            sb.AppendLine($"  【紫微大限警示】宮干 {warnStem} 化忌（{warnSiHua.ji}）入 {dyJiPal}：{jiNote}");
                        }
                    }
                }
                sb.AppendLine();
                // 大限宮位主星格局 + 四化（紫微大限宮位的宮干四化，非八字大運天干）
                // 八字大運可能橫跨多個紫微大限，逐段列出
                if (hasZiwei)
                {
                    var overlapPalaces = DyGetOverlappingDecadePalaces(palaces, lc.startAge, lc.endAge);
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
                        // 主星格局描述（三方四正廟旺）
                        string palaceStars = KbGetPalaceStars(palaces, lcDecadePalace);
                        string contentKey = KbContentPalaceName(lcDecadePalace);
                        string palaceContent = KbFilterZiweiContent(
                            KbExtractPalaceSection(ziweiFullContent, contentKey),
                            KbGetPalaceStarsSet(palaces, lcDecadePalace), chartStars);
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
                            // 有主星：正常顯示主星 + 格局
                            sb.AppendLine($"  宮位主星：{palaceStars}");
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
                string stemCls = flStemElem == jiShenElem ? "X（大忌）"
                    : (flStemElem == yongShenElem || flStemElem == fuYiElem) ? "○（喜用）" : "△（中性）";
                string brCls = string.IsNullOrEmpty(flBrElem) ? "△（中性）"
                    : flBrElem == jiShenElem ? "X（大忌）"
                    : (flBrElem == yongShenElem || flBrElem == fuYiElem) ? "○（喜用）" : "△（中性）";
                sb.AppendLine($"  流年天干 {d.flStem}（{flStemElem}·{d.flStemSS}）：{stemCls}");
                if (!string.IsNullOrEmpty(flBrElem))
                    sb.AppendLine($"  流年地支 {d.flBranch}（{flBrElem}·{d.flBranchSS}）：{brCls}");
                string brEvents = LfBranchEvents(d.flBranch, branches);
                if (!string.IsNullOrEmpty(brEvents))
                    sb.AppendLine($"  歲君互動：{brEvents}");
                sb.AppendLine();

                // 紫微面向
                sb.AppendLine("  ▍ 紫微面向");
                if (YearStemSiHuaMap.TryGetValue(d.flStem, out var siHua))
                {
                    string[] flShTypes = { "化祿", "化權", "化科", "化忌" };
                    string[] flStars   = { siHua.lu, siHua.quan, siHua.ke, siHua.ji };
                    if (hasZiwei && siHuaDescMap.TryGetValue(d.flStem, out var flDyMap))
                    {
                        for (int si = 0; si < flShTypes.Length; si++)
                        {
                            var (pal, desc) = flDyMap.GetValueOrDefault(flShTypes[si], ("", ""));
                            if (!string.IsNullOrEmpty(pal) && LfShouldSkipPalace(pal, d.age)) continue;
                            string palLabel = string.IsNullOrEmpty(pal) ? "（命盤未含此星）" : $"入{pal}";
                            string descText = string.IsNullOrEmpty(desc) ? "" : $"：{desc}";
                            sb.AppendLine($"  {flShTypes[si]}（{flStars[si]}星）{palLabel}{descText}");
                        }
                    }
                    else if (hasZiwei)
                    {
                        for (int si = 0; si < flShTypes.Length; si++)
                        {
                            string pal = KbGetSiHuaPalace(d.flStem, flShTypes[si], palaces);
                            if (!string.IsNullOrEmpty(pal) && LfShouldSkipPalace(pal, d.age)) continue;
                            sb.AppendLine($"  {flShTypes[si]}（{flStars[si]}星）{(string.IsNullOrEmpty(pal) ? "（命盤未含此星）" : "入" + pal)}");
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
            sb.AppendLine("-----------------------------------------------------------------");
            sb.AppendLine("命理大師：玉洞子 | 大運命書 v1.1");
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

            const int cost = 100;
            decimal lnDiscountRate = await GetSubscriptionDiscountRate(user.Id, "book");
            int lnEffectiveCost = (int)Math.Ceiling(cost * lnDiscountRate);
            if (user.Points < lnEffectiveCost)
                return BadRequest(new { error = $"點數不足，需要 {lnEffectiveCost} 點" });

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

                user.Points -= lnEffectiveCost;
                await _context.SaveChangesAsync();
                return Ok(new { result = report, annualSummary, monthlyForecasts, baziTable, luckCycles = scoredCycles, remainingPoints = user.Points });
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
            sb.AppendLine("-----------------------------------------------------------------");
            sb.AppendLine($"命理大師：玉洞子 | 流年命書 v1.0 | {year} 年");
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

            const int cost = 20;
            decimal topicDiscountRate = await GetSubscriptionDiscountRate(user.Id, "consultation");
            int topicEffectiveCost = (int)Math.Ceiling(cost * topicDiscountRate);
            if (user.Points < topicEffectiveCost)
                return BadRequest(new { error = $"點數不足，需要 {topicEffectiveCost} 點" });

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

                user.Points -= topicEffectiveCost;
                await _context.SaveChangesAsync();

                return Ok(new { result = aiResult, remainingPoints = user.Points });
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
        private async Task<decimal> GetSubscriptionDiscountRate(string userId, string productType)
        {
            var now = DateTime.UtcNow;
            var sub = await _context.UserSubscriptions
                .Where(s => s.UserId == userId && s.Status == "active" && s.ExpiryDate > now)
                .OrderByDescending(s => s.ExpiryDate)
                .Include(s => s.Plan)
                .ThenInclude(p => p.Benefits)
                .FirstOrDefaultAsync();

            if (sub == null) return 1.0m;

            var discount = sub.Plan.Benefits.FirstOrDefault(b =>
                b.BenefitType == "discount" && b.ProductType == productType);

            if (discount != null && decimal.TryParse(discount.BenefitValue, out var rate))
                return rate;

            return 1.0m;
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