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
                "問事" => 10,
                _ => 50
            };

            if (user.Points < cost) return BadRequest(new { error = $"點數不足，此功能需要 {cost} 點" });

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

                user.Points -= cost;
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

        private string BuildTopicPrompt(string chartJson, string topic) => $@"
你是一位精通《三命通會》與《紫微斗數》的命理鑑定大師『玉洞子』。
請根據以下 JSON 命盤數據，專門針對命主的【{topic}】問題進行深度鑑定，
給出明確、直接、可執行的命理建議。

### 命盤數據：
{chartJson}

### 問事主題：{topic}

### 輸出格式：

一、{topic} 命格基礎
【命盤顯示】：(此命盤在{topic}方面的先天條件，從八字與紫微雙角度分析)
【喜忌分析】：(哪些五行/星曜對{topic}有利，哪些有害)

二、{topic} 深度研判
【吉象】：(命盤中有利於{topic}的具體星曜/干支，明確說明)
【凶象】：(命盤中不利於{topic}的阻礙，明確說明)
【綜合評斷】：(用一句話斷死{topic}的整體格局，嚴禁使用「可能」或「或許」)

三、{topic} 時機研判
【最佳行動年份】：(從當前大運流年分析，哪幾年最適合)
【需要迴避時段】：(哪幾年不宜輕舉妄動)

四、{topic} 具體行動建議
【立即可做】：(現在就能執行的三件事)
【長期規劃】：(未來 3-5 年的策略方向)
【開運化煞】：(針對{topic}的具體風水或擇日建議)

-----------------------------------------------------------------
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
            if (user.Points < cost)
                return BadRequest(new { error = $"點數不足，需要 {cost} 點" });

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

                // 紫微所在宮位地支（ziwei_patterns_144 查詢基準）
                string ziweiPos     = KbGetZiweiPosition(palaces);

                // ziwei_patterns_144（各宮格局，以紫微地支+目標宮地支查詢）
                string ziweiMing    = await KbZiweiQuery(palaces, "命宮",  ziweiPos);
                string ziweiOffStar = KbGetPalaceStars(palaces, "官祿");
                string ziweiOff     = await KbZiweiQuery(palaces, "官祿",  ziweiPos);
                string ziweiWltStar = KbGetPalaceStars(palaces, "財帛");
                string ziweiWlt     = await KbZiweiQuery(palaces, "財帛",  ziweiPos);
                string ziweiSpsStar = KbGetPalaceStars(palaces, "夫妻");
                string ziweiSps     = await KbZiweiQuery(palaces, "夫妻",  ziweiPos);
                string ziweiHltStar = KbGetPalaceStars(palaces, "疾厄");
                string ziweiHlt     = await KbZiweiQuery(palaces, "疾厄",  ziweiPos);
                string ziweiParStar = KbGetPalaceStars(palaces, "父母");
                string ziweiPar     = await KbZiweiQuery(palaces, "父母",  ziweiPos);
                string ziweiCldStar = KbGetPalaceStars(palaces, "子女");
                string ziweiCld     = await KbZiweiQuery(palaces, "子女",  ziweiPos);
                string ziweiLuck    = string.IsNullOrEmpty(daYunPalace) ? "" : await KbZiweiQuery(palaces, daYunPalace, ziweiPos);

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
                var (mingLuPalace,  mingLuContent)  = await KbGongWeiSiHuaQuery(palaces, "命宮", "化祿");
                var (mingJiPalace,  mingJiContent)  = await KbGongWeiSiHuaQuery(palaces, "命宮", "化忌");
                var (offLuPalace,   offLuContent)   = await KbGongWeiSiHuaQuery(palaces, "官祿宮", "化祿");
                var (offJiPalace,   offJiContent)   = await KbGongWeiSiHuaQuery(palaces, "官祿宮", "化忌");
                var (wltLuPalace,   wltLuContent)   = await KbGongWeiSiHuaQuery(palaces, "財帛宮", "化祿");
                var (wltJiPalace,   wltJiContent)   = await KbGongWeiSiHuaQuery(palaces, "財帛宮", "化忌");
                var (spsLuPalace,   spsLuContent)   = await KbGongWeiSiHuaQuery(palaces, "夫妻宮", "化祿");
                var (spsJiPalace,   spsJiContent)   = await KbGongWeiSiHuaQuery(palaces, "夫妻宮", "化忌");
                var (hltJiPalace,   hltJiContent)   = await KbGongWeiSiHuaQuery(palaces, "疾厄宮", "化忌");

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
                if (!string.IsNullOrEmpty(ziweiMing))  sb_out.AppendLine($"【紫微格局綱領·{mingGongStars}】{ziweiMing}");
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
                // 命宮主星性格（12宮星情）
                if (!string.IsNullOrEmpty(ziweiMing))  sb_out.AppendLine($"【命宮主星·{mingGongStars}】{ziweiMing}");
                // 財帛宮主星對財務個性的影響
                if (!string.IsNullOrEmpty(ziweiWlt))   sb_out.AppendLine($"【財帛宮主星·{ziweiWltStar}（財務個性）】{ziweiWlt}");
                // 官祿宮主星對事業個性的影響
                if (!string.IsNullOrEmpty(ziweiOff))   sb_out.AppendLine($"【官祿宮主星·{ziweiOffStar}（事業個性）】{ziweiOff}");
                if (!string.IsNullOrEmpty(astroN))     sb_out.AppendLine($"【詩評·{astroN}】{KbStripHtml(astroM)}");
                if (!string.IsNullOrEmpty(astroHour))  sb_out.AppendLine($"【先天緣性】{KbStripHtml(astroHour)}");
                bool ch3BaziHas  = !string.IsNullOrEmpty(xgfx) || !string.IsNullOrEmpty(rgxx);
                bool ch3ZiweiHas = !string.IsNullOrEmpty(ziweiMing);
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
                // 官祿宮12宮星情 + 宮位四化飛出
                if (!string.IsNullOrEmpty(ziweiOff))    sb_out.AppendLine($"【官祿宮·{ziweiOffStar}】{ziweiOff}");
                if (!string.IsNullOrEmpty(offLuContent)) sb_out.AppendLine($"【官祿宮化祿飛{offLuPalace}】{offLuContent}");
                if (!string.IsNullOrEmpty(offJiContent)) sb_out.AppendLine($"【官祿宮化忌飛{offJiPalace}】{offJiContent}");
                // 財帛宮12宮星情 + 宮位四化飛出
                if (!string.IsNullOrEmpty(ziweiWlt))    sb_out.AppendLine($"【財帛宮·{ziweiWltStar}】{ziweiWlt}");
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
                // 夫妻宮12宮星情 + 宮位四化飛出
                if (!string.IsNullOrEmpty(ziweiSps))    sb_out.AppendLine($"【夫妻宮·{ziweiSpsStar}】{ziweiSps}");
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
                // 疾厄宮12宮星情 + 化忌飛出
                if (!string.IsNullOrEmpty(ziweiHlt))    sb_out.AppendLine($"【疾厄宮·{ziweiHltStar}】{ziweiHlt}");
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

                // --- 八、大運流年 ---
                sb_out.AppendLine("=== 八、大運流年 ===");
                sb_out.AppendLine("--- 八字大運論 ---");
                if (!string.IsNullOrEmpty(sanMing)) sb_out.AppendLine($"【三命論會】{KbStripHtml(sanMing)}");
                if (!string.IsNullOrEmpty(daYunStem))
                {
                    string daYunFull = KbExpandLiuShen(daYunSS);
                    sb_out.AppendLine($"【當前大運】{daYunStem}{daYunBranch}（{daYunStart}-{daYunEnd}歲，{daYunFull}）");
                }
                sb_out.AppendLine("--- 紫微大運論 ---");
                if (!string.IsNullOrEmpty(daYunPalace) && !string.IsNullOrEmpty(ziweiLuck))
                    sb_out.AppendLine($"【大運宮位·{daYunPalace}】{ziweiLuck}");
                if (!string.IsNullOrEmpty(siHuaJi) && !string.IsNullOrEmpty(daYunPalace) && siHuaJiPalace == daYunPalace)
                    sb_out.AppendLine($"【先天化忌警示·大運入{daYunPalace}】{siHuaJi}");
                sb_out.AppendLine();
                sb_out.AppendLine("-----------------------------------------------------------------");
                sb_out.AppendLine("命理鑑定大師：玉洞子  |  修身齊家，命在人心。  v3.0");

                user.Points -= cost;
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

        // --- KB Helper: 取宮位主星縮寫（第一顆）---
        private static string KbGetPalaceStars(JsonElement palaces, string palaceName)
        {
            if (palaces.ValueKind != JsonValueKind.Array) return "";
            foreach (var p in palaces.EnumerateArray())
            {
                string pname = p.TryGetProperty("palaceName", out var pn) ? pn.GetString() ?? "" :
                               p.TryGetProperty("name", out var n2) ? n2.GetString() ?? "" : "";
                if (pname != palaceName) continue;
                string majorKey = p.TryGetProperty("majorStars", out _) ? "majorStars" : "mainStars";
                if (p.TryGetProperty(majorKey, out var stars) && stars.ValueKind == JsonValueKind.Array)
                    return string.Join(",", stars.EnumerateArray().Select(s => s.GetString() ?? ""));
            }
            return "";
        }

        // --- KB Helper: 查紫微宮位（ziwei_patterns_144，ziwei_position=紫微所在地支，palace_position=目標宮地支）---
        private async Task<string> KbZiweiQuery(JsonElement palaces, string palaceName, string ziweiPos)
        {
            if (string.IsNullOrEmpty(ziweiPos)) return "";
            string palaceBranch = KbGetPalaceBranch(palaces, palaceName);
            if (string.IsNullOrEmpty(palaceBranch)) return "";
            return await KbQuery($"SELECT COALESCE(content,'') AS \"Value\" FROM public.ziwei_patterns_144 WHERE ziwei_position='{ziweiPos}' AND palace_position='{palaceBranch}' LIMIT 1");
        }

        // --- KB Helper: 取宮位地支 ---
        private static string KbGetPalaceBranch(JsonElement palaces, string palaceName)
        {
            if (palaces.ValueKind != JsonValueKind.Array) return "";
            foreach (var p in palaces.EnumerateArray())
            {
                string pname = p.TryGetProperty("palaceName", out var pn) ? pn.GetString() ?? "" :
                               p.TryGetProperty("name", out var n2) ? n2.GetString() ?? "" : "";
                if (pname != palaceName) continue;
                return p.TryGetProperty("earthlyBranch", out var br) ? br.GetString() ?? "" : "";
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
                return p.TryGetProperty("earthlyBranch", out var br) ? br.GetString() ?? "" : "";
            }
            return "";
        }

        // --- KB Helper: 依星曜縮寫找宮位名 ---
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
                        return pname;
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
                if (pname != palaceName) continue;
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
            // Title 格式："命宮四化飛星－－化祿"，ConditionText 格式："命宮化祿入兄弟宮：..."
            string title = $"{sourcePalaceName}四化飛星－－{siHuaType}";
            string condPrefix = $"{sourcePalaceName}{siHuaType}入{targetPalace}";
            string content = await KbQuery($"SELECT COALESCE(\"ResultText\",'') AS \"Value\" FROM \"FortuneRules\" WHERE \"SourceFile\"='宮位四化.docx' AND \"Title\"='{title}' AND \"ConditionText\" LIKE '{condPrefix}%' LIMIT 1");
            return (targetPalace, content);
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
                    return p.TryGetProperty("palaceName", out var pn) ? pn.GetString() ?? "" : "";
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