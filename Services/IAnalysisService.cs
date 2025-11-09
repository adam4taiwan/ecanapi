using Ecanapi.Models.Analysis;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Ecanapi.Services
{
    public interface IAnalysisService
    {
        // --- StarStyle Methods ---
        // 【修正】讓 GetAllStarStylesAsync 可以接收 position 和 mainstar 兩個可選的查詢參數
        Task<IEnumerable<StarStyle>> GetAllStarStylesAsync(float? position, string? mainstar);
        //Task<IEnumerable<StarStyle>> GetAllStarStylesAsync();
        Task<StarStyle?> GetStarStyleByIdAsync(int id);
        Task<StarStyle> CreateStarStyleAsync(StarStyle starStyle);
        Task UpdateStarStyleAsync(StarStyle starStyle);
        Task<bool> DeleteStarStyleAsync(int id);

        // --- 新增的查詢方法 ---
        Task<IEnumerable<PreNatalFourTransformations>> GetPreNatalFourTransformationsAsync(string? mainstar, int? position);
        Task<IEnumerable<PalaceTransformations>> GetPalaceTransformationsAsync(string? mainstar, int? position);
        Task<IEnumerable<EarthlyBranchStars>> GetEarthlyBranchStarsAsync(string? kind, string? skyno, string? toflo);
        Task<IEnumerable<HeavenlyStemStars>> GetHeavenlyStemStarsAsync(string? kind, string? skyno, string? toflo);
        Task<IEnumerable<DayHourStars>> GetDayHourStarsAsync(string? skyFloor, string? position);
        Task<IEnumerable<DayStemToBranch>> GetDayStemToBranchAsync(string? kind, string? skyno, string? toflo);
        Task<IEnumerable<SixtyJiaziDayToHour>> GetSixtyJiaziDayToHourAsync(string? sky, string? month, string? time);
        Task<IEnumerable<DayPillarToMonthBranch>> GetAllDayPillarToMonthBranchesAsync(string? skyFloor, string? position);
        Task<IEnumerable<IChing64Hexagrams>> GetAllIChing64HexagramsAsync(int? guaValue, string? guaName);
        Task<IEnumerable<IChingExplanation>> GetAllIChingExplanationsAsync(int? guaValue, string? guaName);

        // --- PalaceMainStar Methods ---
        Task<IEnumerable<PalaceMainStar>> GetAllPalaceMainStarsAsync();
        Task<PalaceMainStar?> GetPalaceMainStarByIdAsync(int id);
        Task<PalaceMainStar> CreatePalaceMainStarAsync(PalaceMainStar palaceMainStar);
        Task UpdatePalaceMainStarAsync(PalaceMainStar palaceMainStar);
        Task<bool> DeletePalaceMainStarAsync(int id);

        // --- PalaceName Methods ---
        Task<IEnumerable<PalaceName>> GetAllPalaceNamesAsync();
        Task<PalaceName?> GetPalaceNameByIdAsync(int id);
        Task<PalaceName> CreatePalaceNameAsync(PalaceName palaceName);
        Task UpdatePalaceNameAsync(PalaceName palaceName);
        Task<bool> DeletePalaceNameAsync(int id);

        // --- PalaceStarBrightness Methods ---
        Task<IEnumerable<PalaceStarBrightness>> GetAllPalaceStarBrightnessesAsync();
        Task<PalaceStarBrightness?> GetPalaceStarBrightnessByIdAsync(int id);
        Task<PalaceStarBrightness> CreatePalaceStarBrightnessAsync(PalaceStarBrightness palaceStarBrightness);
        Task UpdatePalaceStarBrightnessAsync(PalaceStarBrightness palaceStarBrightness);
        Task<bool> DeletePalaceStarBrightnessAsync(int id);

        // --- EarthlyBranchHiddenStem Methods ---
        Task<IEnumerable<EarthlyBranchHiddenStem>> GetAllEarthlyBranchHiddenStemsAsync();
        Task<EarthlyBranchHiddenStem?> GetEarthlyBranchHiddenStemByIdAsync(int id);
        Task<EarthlyBranchHiddenStem> CreateEarthlyBranchHiddenStemAsync(EarthlyBranchHiddenStem earthlyBranchHiddenStem);
        Task UpdateEarthlyBranchHiddenStemAsync(EarthlyBranchHiddenStem earthlyBranchHiddenStem);
        Task<bool> DeleteEarthlyBranchHiddenStemAsync(int id);

        // --- HeavenlyStemInfo Methods ---
        Task<IEnumerable<HeavenlyStemInfo>> GetAllHeavenlyStemInfosAsync();
        Task<HeavenlyStemInfo?> GetHeavenlyStemInfoByIdAsync(int id);
        Task<HeavenlyStemInfo> CreateHeavenlyStemInfoAsync(HeavenlyStemInfo heavenlyStemInfo);
        Task UpdateHeavenlyStemInfoAsync(HeavenlyStemInfo heavenlyStemInfo);
        Task<bool> DeleteHeavenlyStemInfoAsync(int id);

        // --- NaYin Methods ---
        Task<IEnumerable<NaYin>> GetAllNaYinsAsync();
        Task<NaYin?> GetNaYinByIdAsync(int id);
        Task<NaYin> CreateNaYinAsync(NaYin naYin);
        Task UpdateNaYinAsync(NaYin naYin);
        Task<bool> DeleteNaYinAsync(int id);

        // --- StarCondition Methods ---
        Task<IEnumerable<StarCondition>> GetAllStarConditionsAsync();
        Task<StarCondition?> GetStarConditionByIdAsync(int id);
        Task<StarCondition> CreateStarConditionAsync(StarCondition starCondition);
        Task UpdateStarConditionAsync(StarCondition starCondition);
        Task<bool> DeleteStarConditionAsync(int id);

        // --- BodyMaster Methods ---
        Task<IEnumerable<BodyMaster>> GetAllBodyMastersAsync();
        Task<BodyMaster?> GetBodyMasterByIdAsync(int id);
        Task<BodyMaster> CreateBodyMasterAsync(BodyMaster bodyMaster);
        Task UpdateBodyMasterAsync(BodyMaster bodyMaster);
        Task<bool> DeleteBodyMasterAsync(int id);

        // --- WealthOfficialGeneral Methods ---
        Task<IEnumerable<WealthOfficialGeneral>> GetAllWealthOfficialGeneralsAsync();
        Task<WealthOfficialGeneral?> GetWealthOfficialGeneralByIdAsync(int id);
        Task<WealthOfficialGeneral> CreateWealthOfficialGeneralAsync(WealthOfficialGeneral wealthOfficialGeneral);
        Task UpdateWealthOfficialGeneralAsync(WealthOfficialGeneral wealthOfficialGeneral);
        Task<bool> DeleteWealthOfficialGeneralAsync(int id);

        // --- DayPillarToMonthBranch Methods ---
        Task<IEnumerable<DayPillarToMonthBranch>> GetAllDayPillarToMonthBranchesAsync();
        Task<DayPillarToMonthBranch?> GetDayPillarToMonthBranchByIdAsync(int id);
        Task<DayPillarToMonthBranch> CreateDayPillarToMonthBranchAsync(DayPillarToMonthBranch dayPillarToMonthBranch);
        Task UpdateDayPillarToMonthBranchAsync(DayPillarToMonthBranch dayPillarToMonthBranch);
        Task<bool> DeleteDayPillarToMonthBranchAsync(int id);

        // --- IChing64Hexagrams Methods ---
        Task<IEnumerable<IChing64Hexagrams>> GetAllIChing64HexagramsAsync();
        Task<IChing64Hexagrams?> GetIChing64HexagramsByIdAsync(int id);
        Task<IChing64Hexagrams> CreateIChing64HexagramsAsync(IChing64Hexagrams iChing64Hexagrams);
        Task UpdateIChing64HexagramsAsync(IChing64Hexagrams iChing64Hexagrams);
        Task<bool> DeleteIChing64HexagramsAsync(int id);

        // --- IChingExplanation Methods ---
        Task<IEnumerable<IChingExplanation>> GetAllIChingExplanationsAsync();
        Task<IChingExplanation?> GetIChingExplanationByIdAsync(int id);
        Task<IChingExplanation> CreateIChingExplanationAsync(IChingExplanation iChingExplanation);
        Task UpdateIChingExplanationAsync(IChingExplanation iChingExplanation);
        Task<bool> DeleteIChingExplanationAsync(int id);

        // 【新增】通用的唯讀 SQL 查詢方法
        Task<string> ExecuteRawQueryAsync(string sqlQuery);
        Task<string> ExecuteRawQueryListAsync(string sqlQuery);

    }
}