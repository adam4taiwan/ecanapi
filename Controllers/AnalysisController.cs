using Ecanapi.Models.Analysis;
using Ecanapi.Services;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Ecanapi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class AnalysisController : ControllerBase
    {
        private readonly IAnalysisService _analysisService;

        public AnalysisController(IAnalysisService analysisService)
        {
            _analysisService = analysisService;
        }

        #region --- StarStyle (星曜風格) Endpoints ---
        [HttpGet("starstyle")]
        public async Task<ActionResult<IEnumerable<StarStyle>>> GetStarStyles([FromQuery] float? position, [FromQuery] string? mainstar)
        {
            // 將接收到的參數，傳遞給更新後的服務方法
            var result = await _analysisService.GetAllStarStylesAsync(position, mainstar);
            return Ok(result);
        }
        [HttpGet("starstyle/{id}")]
        public async Task<ActionResult<StarStyle>> GetStarStyle(int id)
        {
            var data = await _analysisService.GetStarStyleByIdAsync(id);
            return data == null ? NotFound() : Ok(data);
        }

        [HttpPost("starstyle")]
        public async Task<ActionResult<StarStyle>> PostStarStyle(StarStyle data)
        {
            var created = await _analysisService.CreateStarStyleAsync(data);
            return CreatedAtAction(nameof(GetStarStyle), new { id = created.UniqueId }, created);
        }

        [HttpPut("starstyle/{id}")]
        public async Task<IActionResult> PutStarStyle(int id, StarStyle data)
        {
            if (id != data.UniqueId) return BadRequest();
            await _analysisService.UpdateStarStyleAsync(data);
            return NoContent();
        }

        [HttpDelete("starstyle/{id}")]
        public async Task<IActionResult> DeleteStarStyle(int id)
        {
            var result = await _analysisService.DeleteStarStyleAsync(id);
            return result ? NoContent() : NotFound();
        }
        #endregion


        #region --- 新增的查詢端點 ---

        [HttpGet("prenatalfourtransformations")]
        public async Task<ActionResult<IEnumerable<PreNatalFourTransformations>>> GetPreNatalFourTransformations([FromQuery] string? mainstar, [FromQuery] int? position)
            => Ok(await _analysisService.GetPreNatalFourTransformationsAsync(mainstar, position));

        [HttpGet("palacetransformations")]
        public async Task<ActionResult<IEnumerable<PalaceTransformations>>> GetPalaceTransformations([FromQuery] string? mainstar, [FromQuery] int? position)
            => Ok(await _analysisService.GetPalaceTransformationsAsync(mainstar, position));

        [HttpGet("earthlybranchstars")]
        public async Task<ActionResult<IEnumerable<EarthlyBranchStars>>> GetEarthlyBranchStars([FromQuery] string? kind, [FromQuery] string? skyno, [FromQuery] string? toflo)
            => Ok(await _analysisService.GetEarthlyBranchStarsAsync(kind, skyno, toflo));

        [HttpGet("heavenlystemstars")]
        public async Task<ActionResult<IEnumerable<HeavenlyStemStars>>> GetHeavenlyStemStars([FromQuery] string? kind, [FromQuery] string? skyno, [FromQuery] string? toflo)
            => Ok(await _analysisService.GetHeavenlyStemStarsAsync(kind, skyno, toflo));

        [HttpGet("dayhourstars")]
        public async Task<ActionResult<IEnumerable<DayHourStars>>> GetDayHourStars([FromQuery] string? skyFloor, [FromQuery] string? position)
            => Ok(await _analysisService.GetDayHourStarsAsync(skyFloor, position));

        [HttpGet("daystemtobranch")]
        public async Task<ActionResult<IEnumerable<DayStemToBranch>>> GetDayStemToBranch([FromQuery] string? kind, [FromQuery] string? skyno, [FromQuery] string? toflo)
            => Ok(await _analysisService.GetDayStemToBranchAsync(kind, skyno, toflo));

        [HttpGet("sixtyjiazidaytohour")]
        public async Task<ActionResult<IEnumerable<SixtyJiaziDayToHour>>> GetSixtyJiaziDayToHour([FromQuery] string? sky, [FromQuery] string? month, [FromQuery] string? time)
            => Ok(await _analysisService.GetSixtyJiaziDayToHourAsync(sky, month, time));

        #endregion

        #region --- 更新的查詢端點 ---
        [HttpGet("daypillartomonthbranch")]
        public async Task<ActionResult<IEnumerable<DayPillarToMonthBranch>>> GetDayPillarToMonthBranches([FromQuery] string? skyFloor, [FromQuery] string? position)
            => Ok(await _analysisService.GetAllDayPillarToMonthBranchesAsync(skyFloor, position));

        [HttpGet("iching64hexagrams")]
        public async Task<ActionResult<IEnumerable<IChing64Hexagrams>>> GetIChing64Hexagrams([FromQuery] int? guaValue, [FromQuery] string? guaName)
            => Ok(await _analysisService.GetAllIChing64HexagramsAsync(guaValue, guaName));

        //[HttpGet("ichingexplanation")]
        //public async Task<ActionResult<IEnumerable<IChingExplanation>>> GetIChingExplanations([FromQuery] int? guaValue, [FromQuery] string? guaName)
        //    => Ok(await _analysisService.GetAllIChingExplanationsAsync(guaValue, guaName));
        #endregion

        #region --- IChingExplanation (易經六十四卦分類解說) ---
        [HttpGet("ichingexplanation")]
        public async Task<ActionResult<IEnumerable<IChingExplanation>>> GetIChingExplanations([FromQuery] int? guaValue, [FromQuery] string? guaName)
            => Ok(await _analysisService.GetAllIChingExplanationsAsync(guaValue, guaName));
        // ... (其他 IChingExplanation 端點維持不變)
        #endregion

        #region --- PalaceMainStar (命宮主星) Endpoints ---
        [HttpGet("palacemainstar")]
        public async Task<ActionResult<IEnumerable<PalaceMainStar>>> GetPalaceMainStars() => Ok(await _analysisService.GetAllPalaceMainStarsAsync());

        [HttpGet("palacemainstar/{id}")]
        public async Task<ActionResult<PalaceMainStar>> GetPalaceMainStar(int id)
        {
            var data = await _analysisService.GetPalaceMainStarByIdAsync(id);
            return data == null ? NotFound() : Ok(data);
        }

        [HttpPost("palacemainstar")]
        public async Task<ActionResult<PalaceMainStar>> PostPalaceMainStar(PalaceMainStar data)
        {
            var created = await _analysisService.CreatePalaceMainStarAsync(data);
            return CreatedAtAction(nameof(GetPalaceMainStar), new { id = created.UniqueId }, created);
        }

        [HttpPut("palacemainstar/{id}")]
        public async Task<IActionResult> PutPalaceMainStar(int id, PalaceMainStar data)
        {
            if (id != data.UniqueId) return BadRequest();
            await _analysisService.UpdatePalaceMainStarAsync(data);
            return NoContent();
        }

        [HttpDelete("palacemainstar/{id}")]
        public async Task<IActionResult> DeletePalaceMainStar(int id)
        {
            var result = await _analysisService.DeletePalaceMainStarAsync(id);
            return result ? NoContent() : NotFound();
        }
        #endregion

        #region --- PalaceName (十二宮稱呼) Endpoints ---
        [HttpGet("palacename")]
        public async Task<ActionResult<IEnumerable<PalaceName>>> GetPalaceNames() => Ok(await _analysisService.GetAllPalaceNamesAsync());

        [HttpGet("palacename/{id}")]
        public async Task<ActionResult<PalaceName>> GetPalaceName(int id)
        {
            var data = await _analysisService.GetPalaceNameByIdAsync(id);
            return data == null ? NotFound() : Ok(data);
        }

        [HttpPost("palacename")]
        public async Task<ActionResult<PalaceName>> PostPalaceName(PalaceName data)
        {
            var created = await _analysisService.CreatePalaceNameAsync(data);
            return CreatedAtAction(nameof(GetPalaceName), new { id = created.UniqueId }, created);
        }

        [HttpPut("palacename/{id}")]
        public async Task<IActionResult> PutPalaceName(int id, PalaceName data)
        {
            if (id != data.UniqueId) return BadRequest();
            await _analysisService.UpdatePalaceNameAsync(data);
            return NoContent();
        }

        [HttpDelete("palacename/{id}")]
        public async Task<IActionResult> DeletePalaceName(int id)
        {
            var result = await _analysisService.DeletePalaceNameAsync(id);
            return result ? NoContent() : NotFound();
        }
        #endregion

        #region --- PalaceStarBrightness (十二宮廟旺) Endpoints ---
        [HttpGet("palacestarbrightness")]
        public async Task<ActionResult<IEnumerable<PalaceStarBrightness>>> GetPalaceStarBrightnesses() => Ok(await _analysisService.GetAllPalaceStarBrightnessesAsync());

        [HttpGet("palacestarbrightness/{id}")]
        public async Task<ActionResult<PalaceStarBrightness>> GetPalaceStarBrightness(int id)
        {
            var data = await _analysisService.GetPalaceStarBrightnessByIdAsync(id);
            return data == null ? NotFound() : Ok(data);
        }

        [HttpPost("palacestarbrightness")]
        public async Task<ActionResult<PalaceStarBrightness>> PostPalaceStarBrightness(PalaceStarBrightness data)
        {
            var created = await _analysisService.CreatePalaceStarBrightnessAsync(data);
            return CreatedAtAction(nameof(GetPalaceStarBrightness), new { id = created.UniqueId }, created);
        }

        [HttpPut("palacestarbrightness/{id}")]
        public async Task<IActionResult> PutPalaceStarBrightness(int id, PalaceStarBrightness data)
        {
            if (id != data.UniqueId) return BadRequest();
            await _analysisService.UpdatePalaceStarBrightnessAsync(data);
            return NoContent();
        }

        [HttpDelete("palacestarbrightness/{id}")]
        public async Task<IActionResult> DeletePalaceStarBrightness(int id)
        {
            var result = await _analysisService.DeletePalaceStarBrightnessAsync(id);
            return result ? NoContent() : NotFound();
        }
        #endregion

        #region --- EarthlyBranchHiddenStem (地支藏干) Endpoints ---
        [HttpGet("earthlybranchhiddenstem")]
        public async Task<ActionResult<IEnumerable<EarthlyBranchHiddenStem>>> GetEarthlyBranchHiddenStems() => Ok(await _analysisService.GetAllEarthlyBranchHiddenStemsAsync());

        [HttpGet("earthlybranchhiddenstem/{id}")]
        public async Task<ActionResult<EarthlyBranchHiddenStem>> GetEarthlyBranchHiddenStem(int id)
        {
            var data = await _analysisService.GetEarthlyBranchHiddenStemByIdAsync(id);
            return data == null ? NotFound() : Ok(data);
        }

        [HttpPost("earthlybranchhiddenstem")]
        public async Task<ActionResult<EarthlyBranchHiddenStem>> PostEarthlyBranchHiddenStem(EarthlyBranchHiddenStem data)
        {
            var created = await _analysisService.CreateEarthlyBranchHiddenStemAsync(data);
            return CreatedAtAction(nameof(GetEarthlyBranchHiddenStem), new { id = created.UniqueId }, created);
        }

        [HttpPut("earthlybranchhiddenstem/{id}")]
        public async Task<IActionResult> PutEarthlyBranchHiddenStem(int id, EarthlyBranchHiddenStem data)
        {
            if (id != data.UniqueId) return BadRequest();
            await _analysisService.UpdateEarthlyBranchHiddenStemAsync(data);
            return NoContent();
        }

        [HttpDelete("earthlybranchhiddenstem/{id}")]
        public async Task<IActionResult> DeleteEarthlyBranchHiddenStem(int id)
        {
            var result = await _analysisService.DeleteEarthlyBranchHiddenStemAsync(id);
            return result ? NoContent() : NotFound();
        }
        #endregion

        #region --- HeavenlyStemInfo (天干陰陽五行) Endpoints ---
        [HttpGet("heavenlysteminfo")]
        public async Task<ActionResult<IEnumerable<HeavenlyStemInfo>>> GetHeavenlyStemInfos() => Ok(await _analysisService.GetAllHeavenlyStemInfosAsync());

        [HttpGet("heavenlysteminfo/{id}")]
        public async Task<ActionResult<HeavenlyStemInfo>> GetHeavenlyStemInfo(int id)
        {
            var data = await _analysisService.GetHeavenlyStemInfoByIdAsync(id);
            return data == null ? NotFound() : Ok(data);
        }

        [HttpPost("heavenlysteminfo")]
        public async Task<ActionResult<HeavenlyStemInfo>> PostHeavenlyStemInfo(HeavenlyStemInfo data)
        {
            var created = await _analysisService.CreateHeavenlyStemInfoAsync(data);
            return CreatedAtAction(nameof(GetHeavenlyStemInfo), new { id = created.UniqueId }, created);
        }

        [HttpPut("heavenlysteminfo/{id}")]
        public async Task<IActionResult> PutHeavenlyStemInfo(int id, HeavenlyStemInfo data)
        {
            if (id != data.UniqueId) return BadRequest();
            await _analysisService.UpdateHeavenlyStemInfoAsync(data);
            return NoContent();
        }

        [HttpDelete("heavenlysteminfo/{id}")]
        public async Task<IActionResult> DeleteHeavenlyStemInfo(int id)
        {
            var result = await _analysisService.DeleteHeavenlyStemInfoAsync(id);
            return result ? NoContent() : NotFound();
        }
        #endregion

        #region --- NaYin (納音) Endpoints ---
        [HttpGet("nayin")]
        public async Task<ActionResult<IEnumerable<NaYin>>> GetNaYins() => Ok(await _analysisService.GetAllNaYinsAsync());

        [HttpGet("nayin/{id}")]
        public async Task<ActionResult<NaYin>> GetNaYin(int id)
        {
            var data = await _analysisService.GetNaYinByIdAsync(id);
            return data == null ? NotFound() : Ok(data);
        }

        [HttpPost("nayin")]
        public async Task<ActionResult<NaYin>> PostNaYin(NaYin data)
        {
            var created = await _analysisService.CreateNaYinAsync(data);
            return CreatedAtAction(nameof(GetNaYin), new { id = created.UniqueId }, created);
        }

        [HttpPut("nayin/{id}")]
        public async Task<IActionResult> PutNaYin(int id, NaYin data)
        {
            if (id != data.UniqueId) return BadRequest();
            await _analysisService.UpdateNaYinAsync(data);
            return NoContent();
        }

        [HttpDelete("nayin/{id}")]
        public async Task<IActionResult> DeleteNaYin(int id)
        {
            var result = await _analysisService.DeleteNaYinAsync(id);
            return result ? NoContent() : NotFound();
        }
        #endregion

        #region --- StarCondition (星曜狀況) Endpoints ---
        [HttpGet("starcondition")]
        public async Task<ActionResult<IEnumerable<StarCondition>>> GetStarConditions() => Ok(await _analysisService.GetAllStarConditionsAsync());

        [HttpGet("starcondition/{id}")]
        public async Task<ActionResult<StarCondition>> GetStarCondition(int id)
        {
            var data = await _analysisService.GetStarConditionByIdAsync(id);
            return data == null ? NotFound() : Ok(data);
        }

        [HttpPost("starcondition")]
        public async Task<ActionResult<StarCondition>> PostStarCondition(StarCondition data)
        {
            var created = await _analysisService.CreateStarConditionAsync(data);
            return CreatedAtAction(nameof(GetStarCondition), new { id = created.UniqueId }, created);
        }

        [HttpPut("starcondition/{id}")]
        public async Task<IActionResult> PutStarCondition(int id, StarCondition data)
        {
            if (id != data.UniqueId) return BadRequest();
            await _analysisService.UpdateStarConditionAsync(data);
            return NoContent();
        }

        [HttpDelete("starcondition/{id}")]
        public async Task<IActionResult> DeleteStarCondition(int id)
        {
            var result = await _analysisService.DeleteStarConditionAsync(id);
            return result ? NoContent() : NotFound();
        }
        #endregion

        #region --- BodyMaster (身主) Endpoints ---
        [HttpGet("bodymaster")]
        public async Task<ActionResult<IEnumerable<BodyMaster>>> GetBodyMasters() => Ok(await _analysisService.GetAllBodyMastersAsync());

        [HttpGet("bodymaster/{id}")]
        public async Task<ActionResult<BodyMaster>> GetBodyMaster(int id)
        {
            var data = await _analysisService.GetBodyMasterByIdAsync(id);
            return data == null ? NotFound() : Ok(data);
        }

        [HttpPost("bodymaster")]
        public async Task<ActionResult<BodyMaster>> PostBodyMaster(BodyMaster data)
        {
            var created = await _analysisService.CreateBodyMasterAsync(data);
            return CreatedAtAction(nameof(GetBodyMaster), new { id = created.UniqueId }, created);
        }

        [HttpPut("bodymaster/{id}")]
        public async Task<IActionResult> PutBodyMaster(int id, BodyMaster data)
        {
            if (id != data.UniqueId) return BadRequest();
            await _analysisService.UpdateBodyMasterAsync(data);
            return NoContent();
        }

        [HttpDelete("bodymaster/{id}")]
        public async Task<IActionResult> DeleteBodyMaster(int id)
        {
            var result = await _analysisService.DeleteBodyMasterAsync(id);
            return result ? NoContent() : NotFound();
        }
        #endregion

        #region --- WealthOfficialGeneral (財官總論) Endpoints ---
        [HttpGet("wealthofficialgeneral")]
        public async Task<ActionResult<IEnumerable<WealthOfficialGeneral>>> GetWealthOfficialGenerals() => Ok(await _analysisService.GetAllWealthOfficialGeneralsAsync());

        [HttpGet("wealthofficialgeneral/{id}")]
        public async Task<ActionResult<WealthOfficialGeneral>> GetWealthOfficialGeneral(int id)
        {
            var data = await _analysisService.GetWealthOfficialGeneralByIdAsync(id);
            return data == null ? NotFound() : Ok(data);
        }

        [HttpPost("wealthofficialgeneral")]
        public async Task<ActionResult<WealthOfficialGeneral>> PostWealthOfficialGeneral(WealthOfficialGeneral data)
        {
            var created = await _analysisService.CreateWealthOfficialGeneralAsync(data);
            return CreatedAtAction(nameof(GetWealthOfficialGeneral), new { id = created.UniqueId }, created);
        }

        [HttpPut("wealthofficialgeneral/{id}")]
        public async Task<IActionResult> PutWealthOfficialGeneral(int id, WealthOfficialGeneral data)
        {
            if (id != data.UniqueId) return BadRequest();
            await _analysisService.UpdateWealthOfficialGeneralAsync(data);
            return NoContent();
        }

        [HttpDelete("wealthofficialgeneral/{id}")]
        public async Task<IActionResult> DeleteWealthOfficialGeneral(int id)
        {
            var result = await _analysisService.DeleteWealthOfficialGeneralAsync(id);
            return result ? NoContent() : NotFound();
        }
        #endregion

        #region --- DayPillarToMonthBranch (日柱對月支) Endpoints ---
        //[HttpGet("daypillartomonthbranch")]
        //public async Task<ActionResult<IEnumerable<DayPillarToMonthBranch>>> GetDayPillarToMonthBranches() => Ok(await _analysisService.GetAllDayPillarToMonthBranchesAsync());

        [HttpGet("daypillartomonthbranch/{id}")]
        public async Task<ActionResult<DayPillarToMonthBranch>> GetDayPillarToMonthBranch(int id)
        {
            var data = await _analysisService.GetDayPillarToMonthBranchByIdAsync(id);
            return data == null ? NotFound() : Ok(data);
        }

        [HttpPost("daypillartomonthbranch")]
        public async Task<ActionResult<DayPillarToMonthBranch>> PostDayPillarToMonthBranch(DayPillarToMonthBranch data)
        {
            var created = await _analysisService.CreateDayPillarToMonthBranchAsync(data);
            return CreatedAtAction(nameof(GetDayPillarToMonthBranch), new { id = created.UniqueId }, created);
        }

        [HttpPut("daypillartomonthbranch/{id}")]
        public async Task<IActionResult> PutDayPillarToMonthBranch(int id, DayPillarToMonthBranch data)
        {
            if (id != data.UniqueId) return BadRequest();
            await _analysisService.UpdateDayPillarToMonthBranchAsync(data);
            return NoContent();
        }

        [HttpDelete("daypillartomonthbranch/{id}")]
        public async Task<IActionResult> DeleteDayPillarToMonthBranch(int id)
        {
            var result = await _analysisService.DeleteDayPillarToMonthBranchAsync(id);
            return result ? NoContent() : NotFound();
        }
        #endregion

        #region --- IChing64Hexagrams (易經六十四卦) Endpoints ---
        //[HttpGet("iching64hexagrams")]
        //public async Task<ActionResult<IEnumerable<IChing64Hexagrams>>> GetIChing64Hexagrams() => Ok(await _analysisService.GetAllIChing64HexagramsAsync());

        [HttpGet("iching64hexagrams/{id}")]
        public async Task<ActionResult<IChing64Hexagrams>> GetIChing64Hexagrams(int id)
        {
            var data = await _analysisService.GetIChing64HexagramsByIdAsync(id);
            return data == null ? NotFound() : Ok(data);
        }

        [HttpPost("iching64hexagrams")]
        public async Task<ActionResult<IChing64Hexagrams>> PostIChing64Hexagrams(IChing64Hexagrams data)
        {
            var created = await _analysisService.CreateIChing64HexagramsAsync(data);
            return CreatedAtAction(nameof(GetIChing64Hexagrams), new { id = created.GuaId }, created);
        }

        [HttpPut("iching64hexagrams/{id}")]
        public async Task<IActionResult> PutIChing64Hexagrams(int id, IChing64Hexagrams data)
        {
            if (id != data.GuaId) return BadRequest();
            await _analysisService.UpdateIChing64HexagramsAsync(data);
            return NoContent();
        }

        [HttpDelete("iching64hexagrams/{id}")]
        public async Task<IActionResult> DeleteIChing64Hexagrams(int id)
        {
            var result = await _analysisService.DeleteIChing64HexagramsAsync(id);
            return result ? NoContent() : NotFound();
        }
        #endregion

        #region --- IChingExplanation (易經六十四卦分類解說) Endpoints ---
        //[HttpGet("ichingexplanation")]
        //public async Task<ActionResult<IEnumerable<IChingExplanation>>> GetIChingExplanations() => Ok(await _analysisService.GetAllIChingExplanationsAsync());

        [HttpGet("ichingexplanation/{id}")]
        public async Task<ActionResult<IChingExplanation>> GetIChingExplanation(int id)
        {
            var data = await _analysisService.GetIChingExplanationByIdAsync(id);
            return data == null ? NotFound() : Ok(data);
        }

        [HttpPost("ichingexplanation")]
        public async Task<ActionResult<IChingExplanation>> PostIChingExplanation(IChingExplanation data)
        {
            var created = await _analysisService.CreateIChingExplanationAsync(data);
            return CreatedAtAction(nameof(GetIChingExplanation), new { id = created.UniqueId }, created);
        }

        [HttpPut("ichingexplanation/{id}")]
        public async Task<IActionResult> PutIChingExplanation(int id, IChingExplanation data)
        {
            if (id != data.UniqueId) return BadRequest();
            await _analysisService.UpdateIChingExplanationAsync(data);
            return NoContent();
        }

        [HttpDelete("ichingexplanation/{id}")]
        public async Task<IActionResult> DeleteIChingExplanation(int id)
        {
            var result = await _analysisService.DeleteIChingExplanationAsync(id);
            return result ? NoContent() : NotFound();
        }
        #endregion
        // 【新增】通用的唯讀 SQL 查詢端點
        [HttpPost("query")]
        public async Task<ActionResult<string>> ExecuteRawQuery([FromBody] RawSqlQueryRequest request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Query))
            {
                return BadRequest("SQL query cannot be empty.");
            }

            string result = await _analysisService.ExecuteRawQueryAsync(request.Query);

            return Ok(result);
        }

        // 【新增】一個簡單的資料傳輸物件(DTO)來接收查詢字串
        public record RawSqlQueryRequest(string Query);

        [HttpPost("query/list")]
        public async Task<ActionResult<string>> ExecuteRawQueryList([FromBody] RawSqlQueryRequest request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Query))
            {
                return BadRequest("SQL query cannot be empty.");
            }

            string result = await _analysisService.ExecuteRawQueryListAsync(request.Query);

            return Ok(result);
        }
    }
}