using Ecanapi.Data; // 請確認這是您 DbContext 所在的正確命名空間
using Ecanapi.Models.Analysis;
using Microsoft.EntityFrameworkCore;
using Npgsql;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace Ecanapi.Services
{
    public class AnalysisService : IAnalysisService
    {
        private readonly ApplicationDbContext _context; // 請將 ApplicationDbContext 換成您實際的 DbContext 類別名稱

        public AnalysisService(ApplicationDbContext context) // 同上，請換成您實際的 DbContext 類別名稱
        {
            _context = context;
        }

        // --- StarStyle ---
        public async Task<IEnumerable<StarStyle>> GetAllStarStylesAsync(float? position, string? mainstar)
        {
            // 建立一個可查詢的基礎
            var query = _context.StarStyles.AsQueryable();

            // 如果 position 參數有值，則加入 position 的查詢條件
            if (position.HasValue)
            {
                query = query.Where(s => s.Position == position.Value);
            }

            // 如果 mainstar 參數有值，則加入 mainstar 的查詢條件
            if (!string.IsNullOrEmpty(mainstar))
            {
                // 為了處理 "紫破" 或 "破紫" 的情況，我們先將傳入的字串反轉
                string reversedMainstar = new string(mainstar.Reverse().ToArray());
                query = query.Where(s => s.MainStar != null && (s.MainStar.Contains(mainstar) || s.MainStar.Contains(reversedMainstar)));
            }

            return await query.ToListAsync();
        }
        public async Task<IEnumerable<StarStyle>> GetAllStarStylesAsync() => await _context.StarStyles.ToListAsync();
        public async Task<StarStyle?> GetStarStyleByIdAsync(int id) => await _context.StarStyles.FindAsync(id);
        public async Task<StarStyle> CreateStarStyleAsync(StarStyle data) { _context.StarStyles.Add(data); await _context.SaveChangesAsync(); return data; }
        public async Task UpdateStarStyleAsync(StarStyle data) { _context.Entry(data).State = EntityState.Modified; await _context.SaveChangesAsync(); }
        public async Task<bool> DeleteStarStyleAsync(int id) { var data = await GetStarStyleByIdAsync(id); if (data == null) return false; _context.StarStyles.Remove(data); await _context.SaveChangesAsync(); return true; }

        // --- PalaceMainStar ---
        public async Task<IEnumerable<PalaceMainStar>> GetAllPalaceMainStarsAsync() => await _context.PalaceMainStars.ToListAsync();
        public async Task<PalaceMainStar?> GetPalaceMainStarByIdAsync(int id) => await _context.PalaceMainStars.FindAsync(id);
        public async Task<PalaceMainStar> CreatePalaceMainStarAsync(PalaceMainStar data) { _context.PalaceMainStars.Add(data); await _context.SaveChangesAsync(); return data; }
        public async Task UpdatePalaceMainStarAsync(PalaceMainStar data) { _context.Entry(data).State = EntityState.Modified; await _context.SaveChangesAsync(); }
        public async Task<bool> DeletePalaceMainStarAsync(int id) { var data = await GetPalaceMainStarByIdAsync(id); if (data == null) return false; _context.PalaceMainStars.Remove(data); await _context.SaveChangesAsync(); return true; }

        // --- PalaceName ---
        public async Task<IEnumerable<PalaceName>> GetAllPalaceNamesAsync() => await _context.PalaceNames.ToListAsync();
        public async Task<PalaceName?> GetPalaceNameByIdAsync(int id) => await _context.PalaceNames.FindAsync(id);
        public async Task<PalaceName> CreatePalaceNameAsync(PalaceName data) { _context.PalaceNames.Add(data); await _context.SaveChangesAsync(); return data; }
        public async Task UpdatePalaceNameAsync(PalaceName data) { _context.Entry(data).State = EntityState.Modified; await _context.SaveChangesAsync(); }
        public async Task<bool> DeletePalaceNameAsync(int id) { var data = await GetPalaceNameByIdAsync(id); if (data == null) return false; _context.PalaceNames.Remove(data); await _context.SaveChangesAsync(); return true; }

        // --- PalaceStarBrightness ---
        public async Task<IEnumerable<PalaceStarBrightness>> GetAllPalaceStarBrightnessesAsync() => await _context.PalaceStarBrightnesses.ToListAsync();
        public async Task<PalaceStarBrightness?> GetPalaceStarBrightnessByIdAsync(int id) => await _context.PalaceStarBrightnesses.FindAsync(id);
        public async Task<PalaceStarBrightness> CreatePalaceStarBrightnessAsync(PalaceStarBrightness data) { _context.PalaceStarBrightnesses.Add(data); await _context.SaveChangesAsync(); return data; }
        public async Task UpdatePalaceStarBrightnessAsync(PalaceStarBrightness data) { _context.Entry(data).State = EntityState.Modified; await _context.SaveChangesAsync(); }
        public async Task<bool> DeletePalaceStarBrightnessAsync(int id) { var data = await GetPalaceStarBrightnessByIdAsync(id); if (data == null) return false; _context.PalaceStarBrightnesses.Remove(data); await _context.SaveChangesAsync(); return true; }

        // --- EarthlyBranchHiddenStem ---
        public async Task<IEnumerable<EarthlyBranchHiddenStem>> GetAllEarthlyBranchHiddenStemsAsync() => await _context.EarthlyBranchHiddenStems.ToListAsync();
        public async Task<EarthlyBranchHiddenStem?> GetEarthlyBranchHiddenStemByIdAsync(int id) => await _context.EarthlyBranchHiddenStems.FindAsync(id);
        public async Task<EarthlyBranchHiddenStem> CreateEarthlyBranchHiddenStemAsync(EarthlyBranchHiddenStem data) { _context.EarthlyBranchHiddenStems.Add(data); await _context.SaveChangesAsync(); return data; }
        public async Task UpdateEarthlyBranchHiddenStemAsync(EarthlyBranchHiddenStem data) { _context.Entry(data).State = EntityState.Modified; await _context.SaveChangesAsync(); }
        public async Task<bool> DeleteEarthlyBranchHiddenStemAsync(int id) { var data = await GetEarthlyBranchHiddenStemByIdAsync(id); if (data == null) return false; _context.EarthlyBranchHiddenStems.Remove(data); await _context.SaveChangesAsync(); return true; }

        // --- HeavenlyStemInfo ---
        public async Task<IEnumerable<HeavenlyStemInfo>> GetAllHeavenlyStemInfosAsync() => await _context.HeavenlyStemInfos.ToListAsync();
        public async Task<HeavenlyStemInfo?> GetHeavenlyStemInfoByIdAsync(int id) => await _context.HeavenlyStemInfos.FindAsync(id);
        public async Task<HeavenlyStemInfo> CreateHeavenlyStemInfoAsync(HeavenlyStemInfo data) { _context.HeavenlyStemInfos.Add(data); await _context.SaveChangesAsync(); return data; }
        public async Task UpdateHeavenlyStemInfoAsync(HeavenlyStemInfo data) { _context.Entry(data).State = EntityState.Modified; await _context.SaveChangesAsync(); }
        public async Task<bool> DeleteHeavenlyStemInfoAsync(int id) { var data = await GetHeavenlyStemInfoByIdAsync(id); if (data == null) return false; _context.HeavenlyStemInfos.Remove(data); await _context.SaveChangesAsync(); return true; }

        // --- NaYin ---
        public async Task<IEnumerable<NaYin>> GetAllNaYinsAsync() => await _context.NaYins.ToListAsync();
        public async Task<NaYin?> GetNaYinByIdAsync(int id) => await _context.NaYins.FindAsync(id);
        public async Task<NaYin> CreateNaYinAsync(NaYin data) { _context.NaYins.Add(data); await _context.SaveChangesAsync(); return data; }
        public async Task UpdateNaYinAsync(NaYin data) { _context.Entry(data).State = EntityState.Modified; await _context.SaveChangesAsync(); }
        public async Task<bool> DeleteNaYinAsync(int id) { var data = await GetNaYinByIdAsync(id); if (data == null) return false; _context.NaYins.Remove(data); await _context.SaveChangesAsync(); return true; }

        // --- StarCondition ---
        public async Task<IEnumerable<StarCondition>> GetAllStarConditionsAsync() => await _context.StarConditions.ToListAsync();
        public async Task<StarCondition?> GetStarConditionByIdAsync(int id) => await _context.StarConditions.FindAsync(id);
        public async Task<StarCondition> CreateStarConditionAsync(StarCondition data) { _context.StarConditions.Add(data); await _context.SaveChangesAsync(); return data; }
        public async Task UpdateStarConditionAsync(StarCondition data) { _context.Entry(data).State = EntityState.Modified; await _context.SaveChangesAsync(); }
        public async Task<bool> DeleteStarConditionAsync(int id) { var data = await GetStarConditionByIdAsync(id); if (data == null) return false; _context.StarConditions.Remove(data); await _context.SaveChangesAsync(); return true; }

        // --- BodyMaster ---
        public async Task<IEnumerable<BodyMaster>> GetAllBodyMastersAsync() => await _context.BodyMasters.ToListAsync();
        public async Task<BodyMaster?> GetBodyMasterByIdAsync(int id) => await _context.BodyMasters.FindAsync(id);
        public async Task<BodyMaster> CreateBodyMasterAsync(BodyMaster data) { _context.BodyMasters.Add(data); await _context.SaveChangesAsync(); return data; }
        public async Task UpdateBodyMasterAsync(BodyMaster data) { _context.Entry(data).State = EntityState.Modified; await _context.SaveChangesAsync(); }
        public async Task<bool> DeleteBodyMasterAsync(int id) { var data = await GetBodyMasterByIdAsync(id); if (data == null) return false; _context.BodyMasters.Remove(data); await _context.SaveChangesAsync(); return true; }

        // --- WealthOfficialGeneral ---
        public async Task<IEnumerable<WealthOfficialGeneral>> GetAllWealthOfficialGeneralsAsync() => await _context.WealthOfficialGenerals.ToListAsync();
        public async Task<WealthOfficialGeneral?> GetWealthOfficialGeneralByIdAsync(int id) => await _context.WealthOfficialGenerals.FindAsync(id);
        public async Task<WealthOfficialGeneral> CreateWealthOfficialGeneralAsync(WealthOfficialGeneral data) { _context.WealthOfficialGenerals.Add(data); await _context.SaveChangesAsync(); return data; }
        public async Task UpdateWealthOfficialGeneralAsync(WealthOfficialGeneral data) { _context.Entry(data).State = EntityState.Modified; await _context.SaveChangesAsync(); }
        public async Task<bool> DeleteWealthOfficialGeneralAsync(int id) { var data = await GetWealthOfficialGeneralByIdAsync(id); if (data == null) return false; _context.WealthOfficialGenerals.Remove(data); await _context.SaveChangesAsync(); return true; }

        // --- DayPillarToMonthBranch ---
        public async Task<IEnumerable<DayPillarToMonthBranch>> GetAllDayPillarToMonthBranchesAsync() => await _context.DayPillarToMonthBranches.ToListAsync();
        public async Task<DayPillarToMonthBranch?> GetDayPillarToMonthBranchByIdAsync(int id) => await _context.DayPillarToMonthBranches.FindAsync(id);
        public async Task<DayPillarToMonthBranch> CreateDayPillarToMonthBranchAsync(DayPillarToMonthBranch data) { _context.DayPillarToMonthBranches.Add(data); await _context.SaveChangesAsync(); return data; }
        public async Task UpdateDayPillarToMonthBranchAsync(DayPillarToMonthBranch data) { _context.Entry(data).State = EntityState.Modified; await _context.SaveChangesAsync(); }
        public async Task<bool> DeleteDayPillarToMonthBranchAsync(int id) { var data = await GetDayPillarToMonthBranchByIdAsync(id); if (data == null) return false; _context.DayPillarToMonthBranches.Remove(data); await _context.SaveChangesAsync(); return true; }

        // --- IChing64Hexagrams ---
        public async Task<IEnumerable<IChing64Hexagrams>> GetAllIChing64HexagramsAsync() => await _context.IChing64Hexagrams.ToListAsync();
        public async Task<IChing64Hexagrams?> GetIChing64HexagramsByIdAsync(int id) => await _context.IChing64Hexagrams.FindAsync(id);
        public async Task<IChing64Hexagrams> CreateIChing64HexagramsAsync(IChing64Hexagrams data) { _context.IChing64Hexagrams.Add(data); await _context.SaveChangesAsync(); return data; }
        public async Task UpdateIChing64HexagramsAsync(IChing64Hexagrams data) { _context.Entry(data).State = EntityState.Modified; await _context.SaveChangesAsync(); }
        public async Task<bool> DeleteIChing64HexagramsAsync(int id) { var data = await GetIChing64HexagramsByIdAsync(id); if (data == null) return false; _context.IChing64Hexagrams.Remove(data); await _context.SaveChangesAsync(); return true; }

        // --- IChingExplanation ---
        public async Task<IEnumerable<IChingExplanation>> GetAllIChingExplanationsAsync() => await _context.IChingExplanations.ToListAsync();
        public async Task<IChingExplanation?> GetIChingExplanationByIdAsync(int id) => await _context.IChingExplanations.FindAsync(id);
        public async Task<IChingExplanation> CreateIChingExplanationAsync(IChingExplanation data) { _context.IChingExplanations.Add(data); await _context.SaveChangesAsync(); return data; }
        public async Task UpdateIChingExplanationAsync(IChingExplanation data) { _context.Entry(data).State = EntityState.Modified; await _context.SaveChangesAsync(); }
        public async Task<bool> DeleteIChingExplanationAsync(int id) { var data = await GetIChingExplanationByIdAsync(id); if (data == null) return false; _context.IChingExplanations.Remove(data); await _context.SaveChangesAsync(); return true; }

        // --- 先天四化入十二宮 ---
        public async Task<IEnumerable<PreNatalFourTransformations>> GetPreNatalFourTransformationsAsync(string? mainstar, int? position)
        {
            var query = _context.PreNatalFourTransformations.AsQueryable();
            if (!string.IsNullOrEmpty(mainstar)) query = query.Where(x => x.MainStar == mainstar);
            if (position.HasValue) query = query.Where(x => x.Position == position.Value);
            return await query.ToListAsync();
        }

        // --- 十二宮化入十二宮 ---
        public async Task<IEnumerable<PalaceTransformations>> GetPalaceTransformationsAsync(string? mainstar, int? position)
        {
            var query = _context.PalaceTransformations.AsQueryable();
            if (!string.IsNullOrEmpty(mainstar)) query = query.Where(x => x.MainStar == mainstar);
            if (position.HasValue) query = query.Where(x => x.Position == position.Value);
            return await query.ToListAsync();
        }

        // --- 地支星剎 ---
        // 【修正】使用 FromSqlRaw 繞過 'DESC' 關鍵字問題
        public async Task<IEnumerable<EarthlyBranchStars>> GetEarthlyBranchStarsAsync(string? kind, string? skyno, string? toflo)
        {
            var sql = new StringBuilder("SELECT * FROM \"地支星剎\" WHERE 1=1");
            var parameters = new List<NpgsqlParameter>();

            if (!string.IsNullOrEmpty(kind))
            {
                sql.Append(" AND \"KIND\" = @kind");
                parameters.Add(new NpgsqlParameter("@kind", kind));
            }
            if (!string.IsNullOrEmpty(skyno))
            {
                sql.Append(" AND \"SKYNO\" = @skyno");
                parameters.Add(new NpgsqlParameter("@skyno", skyno));
            }
            if (!string.IsNullOrEmpty(toflo))
            {
                sql.Append(" AND \"TOFLO\" = @toflo");
                parameters.Add(new NpgsqlParameter("@toflo", toflo));
            }
            return await _context.EarthlyBranchStars.FromSqlRaw(sql.ToString(), parameters.ToArray()).ToListAsync();
        }

        // --- 天干星剎 ---
        // 【修正】使用 FromSqlRaw 繞過 'DESC' 關鍵字問題
        public async Task<IEnumerable<HeavenlyStemStars>> GetHeavenlyStemStarsAsync(string? kind, string? skyno, string? toflo)
        {
            var sql = new StringBuilder("SELECT * FROM \"天干星剎\" WHERE 1=1");
            var parameters = new List<NpgsqlParameter>();

            if (!string.IsNullOrEmpty(kind))
            {
                sql.Append(" AND \"KIND\" = @kind");
                parameters.Add(new NpgsqlParameter("@kind", kind));
            }
            if (!string.IsNullOrEmpty(skyno))
            {
                sql.Append(" AND \"SKYNO\" = @skyno");
                parameters.Add(new NpgsqlParameter("@skyno", skyno));
            }
            if (!string.IsNullOrEmpty(toflo))
            {
                sql.Append(" AND \"TOFLO\" = @toflo");
                parameters.Add(new NpgsqlParameter("@toflo", toflo));
            }
            return await _context.HeavenlyStemStars.FromSqlRaw(sql.ToString(), parameters.ToArray()).ToListAsync();
        }

        // --- 日對時星剎 ---
        // 【修正】使用 FromSqlRaw 繞過 'desc' 關鍵字問題
        public async Task<IEnumerable<DayHourStars>> GetDayHourStarsAsync(string? skyFloor, string? position)
        {
            var sql = new StringBuilder("SELECT * FROM \"日對時星剎\" WHERE 1=1");
            var parameters = new List<NpgsqlParameter>();
            if (!string.IsNullOrEmpty(skyFloor))
            {
                sql.Append(" AND \"SkyFloor\" = @skyFloor");
                parameters.Add(new NpgsqlParameter("@skyFloor", skyFloor));
            }
            if (!string.IsNullOrEmpty(position))
            {
                sql.Append(" AND \"position\" = @position");
                parameters.Add(new NpgsqlParameter("@position", position));
            }
            return await _context.DayHourStars.FromSqlRaw(sql.ToString(), parameters.ToArray()).ToListAsync();
        }

        // --- 日干對地支 ---
        // 【修正】使用 FromSqlRaw 繞過 'DESC' 關鍵字問題
        public async Task<IEnumerable<DayStemToBranch>> GetDayStemToBranchAsync(string? kind, string? skyno, string? toflo)
        {
            var sql = new StringBuilder("SELECT * FROM \"日干對地支\" WHERE 1=1");
            var parameters = new List<NpgsqlParameter>();
            if (!string.IsNullOrEmpty(kind))
            {
                sql.Append(" AND \"KIND\" = @kind");
                parameters.Add(new NpgsqlParameter("@kind", kind));
            }
            if (!string.IsNullOrEmpty(skyno))
            {
                sql.Append(" AND \"SKYNO\" = @skyno");
                parameters.Add(new NpgsqlParameter("@skyno", skyno));
            }
            if (!string.IsNullOrEmpty(toflo))
            {
                sql.Append(" AND \"TOFLO\" = @toflo");
                parameters.Add(new NpgsqlParameter("@toflo", toflo));
            }
            return await _context.DayStemToBranches.FromSqlRaw(sql.ToString(), parameters.ToArray()).ToListAsync();
        }

        // --- 六十甲子日對時 ---
        // 【修正】使用 FromSqlRaw 繞過 'desc' 關鍵字問題
        // --- 六十甲子日時查詢 ---
        //public async Task<string> GetSixtyJiaziDayToHourAsync(string sky, string month, string time)
        //{
        //    var dayPillar = $"{sky}{month}";
        //    var sqlQuery = $"SELECT CONCAT(rgcz,xgfx,aqfx,syfx,cyfx,jkfx) AS \"Value\" FROM public.\"六十甲子命主\" WHERE rgz = {dayPillar} AND time = {time}";

        //    try
        //    {
        //        var result = await _context.Database
        //            .SqlQuery<string>(FormattableStringFactory.Create(sqlQuery, dayPillar, time))
        //            .ToListAsync();

        //        if (result.Count == 1)
        //        {
        //            return result[0] ?? "";
        //        }

        //        return "查詢結果多筆或無數據";
        //    }
        //    catch (Exception ex)
        //    {
        //        return $"Error: {ex.Message}";
        //    }
        //}
        public async Task<IEnumerable<SixtyJiaziDayToHour>> GetSixtyJiaziDayToHourAsync(string? sky, string? month, string? time)
        {
            var sql = new StringBuilder("SELECT * FROM \"六十甲子日對時\" WHERE 1=1");
            var parameters = new List<NpgsqlParameter>();
            if (!string.IsNullOrEmpty(sky))
            {
                sql.Append(" AND \"Sky\" = @sky");
                parameters.Add(new NpgsqlParameter("@sky", sky));
            }
            if (!string.IsNullOrEmpty(month))
            {
                sql.Append(" AND \"Month\" = @month");
                parameters.Add(new NpgsqlParameter("@month", month));
            }
            if (!string.IsNullOrEmpty(time))
            {
                sql.Append(" AND \"time\" = @time");
                parameters.Add(new NpgsqlParameter("@time", time));
            }
            return await _context.SixtyJiaziDayToHours.FromSqlRaw(sql.ToString(), parameters.ToArray()).ToListAsync();

        }

        // --- DayPillarToMonthBranch (日柱對月支) ---
        public async Task<IEnumerable<DayPillarToMonthBranch>> GetAllDayPillarToMonthBranchesAsync(string? skyFloor, string? position)
        {
            var sql = new StringBuilder("SELECT * FROM \"日柱對月支\" WHERE 1=1");
            var parameters = new List<NpgsqlParameter>();
            if (!string.IsNullOrEmpty(skyFloor))
            {
                sql.Append(" AND \"SkyFloor\" = @skyFloor");
                parameters.Add(new NpgsqlParameter("@skyFloor", skyFloor));
            }
            if (!string.IsNullOrEmpty(position))
            {
                sql.Append(" AND \"position\" = @position");
                parameters.Add(new NpgsqlParameter("@position", position));
            }
            return await _context.DayPillarToMonthBranches.FromSqlRaw(sql.ToString(), parameters.ToArray()).ToListAsync();
        }
        // ... (DayPillarToMonthBranch 的其他 CRUD 方法維持不變)

        // --- IChing64Hexagrams (易經六十四卦) ---
        public async Task<IEnumerable<IChing64Hexagrams>> GetAllIChing64HexagramsAsync(int? guaValue, string? guaName)
        {
            var query = _context.IChing64Hexagrams.AsQueryable();
            if (guaValue.HasValue) query = query.Where(x => x.GuaValue == guaValue.Value);
            if (!string.IsNullOrEmpty(guaName)) query = query.Where(x => x.GuaName == guaName);
            return await query.ToListAsync();
        }
        // ... (IChing64Hexagrams 的其他 CRUD 方法維持不變)
        // 【新增】通用的唯讀 SQL 查詢方法實作
        //public async Task<string> ExecuteRawQueryAsync(string sqlQuery)
        //{
        //    // 安全性檢查：只允許 SELECT 查詢
        //    if (string.IsNullOrWhiteSpace(sqlQuery) || !sqlQuery.TrimStart().StartsWith("SELECT", StringComparison.OrdinalIgnoreCase))
        //    {
        //        return "Error: Only SELECT queries are allowed.";
        //    }

        //    try
        //    {
        //        await using var connection = _context.Database.GetDbConnection();
        //        await connection.OpenAsync();

        //        await using var command = connection.CreateCommand();
        //        command.CommandText = sqlQuery;

        //        var result = await command.ExecuteScalarAsync();

        //        await connection.CloseAsync();

        //        return result?.ToString() ?? "";
        //    }
        //    catch (Exception ex)
        //    {
        //        // 如果 SQL 執行出錯，回傳錯誤訊息，方便除錯
        //        return $"Error: {ex.Message}";
        //    }
        //}

        // --- Raw Query ---
        public async Task<string> ExecuteRawQueryAsync(string sqlQuery)
        {
            // 安全性檢查：只允許 SELECT 查詢
            if (string.IsNullOrWhiteSpace(sqlQuery) ||
                !sqlQuery.TrimStart().StartsWith("SELECT", StringComparison.OrdinalIgnoreCase))
            {
                return "Error: Only SELECT queries are allowed.";
            }

            try
            {
                var result = await _context.Database
                    .SqlQueryRaw<string>(sqlQuery)
                    .ToListAsync();

                if (result.Count == 1)
                {
                    return result[0] ?? "";
                }

                return JsonSerializer.Serialize(result, new JsonSerializerOptions
                {
                    WriteIndented = true
                });
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }
        //public async Task<string> ExecuteRawQueryAsync(string sqlQuery)
        //{
        //    // 安全性檢查
        //    if (string.IsNullOrWhiteSpace(sqlQuery) || !sqlQuery.TrimStart().StartsWith("SELECT", StringComparison.OrdinalIgnoreCase))
        //        return "Error: Only SELECT queries are allowed.";

        //    try
        //    {
        //        // 用 EF 的 SqlQueryRaw 處理 scalar（單一值），自動管理連接
        //        // 注意：如果 sqlQuery 有參數，這裡暫用無參；下面會教參數化
        //        var result = await _context.Database.SqlQueryRaw<string>(sqlQuery).FirstOrDefaultAsync();
        //        return result ?? "";
        //    }
        //    catch (Exception ex)
        //    {
        //        // 可加 ILogger 記錄
        //        return $"Error: {ex.Message}";
        //    }
        //}
        public async Task<string> ExecuteRawQueryFlexibleAsync(string sqlQuery)
        {
            if (string.IsNullOrWhiteSpace(sqlQuery) ||
                !sqlQuery.TrimStart().StartsWith("SELECT", StringComparison.OrdinalIgnoreCase))
            {
                return "Error: Only SELECT queries are allowed.";
            }

            try
            {
                await using var connection = _context.Database.GetDbConnection();
                await connection.OpenAsync();

                await using var command = connection.CreateCommand();
                command.CommandText = sqlQuery;

                // 嘗試讀取多筆
                var resultList = new List<Dictionary<string, object>>();
                await using var reader = await command.ExecuteReaderAsync();

                while (await reader.ReadAsync())
                {
                    var row = new Dictionary<string, object>();
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        row[reader.GetName(i)] = reader.IsDBNull(i) ? null : reader.GetValue(i);
                    }
                    resultList.Add(row);
                }

                await connection.CloseAsync();

                // 自動判斷：如果只有一筆一欄 → 直接回傳文字
                if (resultList.Count == 1 && resultList[0].Count == 1)
                {
                    return resultList[0].Values.FirstOrDefault()?.ToString() ?? "";
                }

                // 多筆 → JSON
                return JsonSerializer.Serialize(resultList, new JsonSerializerOptions
                {
                    WriteIndented = true
                });
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }

        // QueryList
        // --- QueryList ---
        public async Task<string> ExecuteRawQueryListAsync(string sqlQuery)
        {
            // 安全性檢查：只允許 SELECT 查詢
            if (string.IsNullOrWhiteSpace(sqlQuery) ||
                !sqlQuery.TrimStart().StartsWith("SELECT", StringComparison.OrdinalIgnoreCase))
            {
                return "Error: Only SELECT queries are allowed.";
            }

            try
            {
                // 使用 EF Core 提供的連線，確保由 DbContext 管理
                await using var connection = _context.Database.GetDbConnection();
                await connection.OpenAsync();

                await using var command = connection.CreateCommand();
                command.CommandText = sqlQuery;

                var resultList = new List<Dictionary<string, object>>();

                await using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    var row = new Dictionary<string, object>();
                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        row[reader.GetName(i)] = reader.IsDBNull(i) ? null : reader.GetValue(i);
                    }
                    resultList.Add(row);
                }

                // 連線由 await using 自動關閉，無需手動 CloseAsync

                // 轉成 JSON 字串
                return JsonSerializer.Serialize(resultList, new JsonSerializerOptions
                {
                    WriteIndented = true
                });
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }
        //public async Task<string> ExecuteRawQueryListAsync(string sqlQuery)
        //{
        //    // 安全性檢查：只允許 SELECT 查詢
        //    if (string.IsNullOrWhiteSpace(sqlQuery) ||
        //        !sqlQuery.TrimStart().StartsWith("SELECT", StringComparison.OrdinalIgnoreCase))
        //    {
        //        return "Error: Only SELECT queries are allowed.";
        //    }

        //    try
        //    {
        //        await using var connection = _context.Database.GetDbConnection();
        //        await connection.OpenAsync();

        //        await using var command = connection.CreateCommand();
        //        command.CommandText = sqlQuery;

        //        var resultList = new List<Dictionary<string, object>>();

        //        await using var reader = await command.ExecuteReaderAsync();
        //        while (await reader.ReadAsync())
        //        {
        //            var row = new Dictionary<string, object>();

        //            for (int i = 0; i < reader.FieldCount; i++)
        //            {
        //                row[reader.GetName(i)] = reader.IsDBNull(i) ? null : reader.GetValue(i);
        //            }

        //            resultList.Add(row);
        //        }

        //        await connection.CloseAsync();

        //        // 轉成 JSON 字串
        //        return JsonSerializer.Serialize(resultList, new JsonSerializerOptions
        //        {
        //            WriteIndented = true
        //        });
        //    }
        //    catch (Exception ex)
        //    {
        //        return $"Error: {ex.Message}";
        //    }
        //}


        // --- IChingExplanation (易經六十四卦分類解說) ---
        public async Task<IEnumerable<IChingExplanation>> GetAllIChingExplanationsAsync(int? guaValue, string? guaName)
        {
            var query = _context.IChingExplanations.AsQueryable();

            if (!string.IsNullOrEmpty(guaName))
            {
                var hexagram = await _context.IChing64Hexagrams.FirstOrDefaultAsync(h => h.GuaName == guaName);
                if (hexagram != null)
                {
                    query = query.Where(x => x.GuaId == hexagram.GuaId);
                }
                else
                {
                    return new List<IChingExplanation>(); // Or handle as not found
                }
            }
            else if (guaValue.HasValue)
            {
                query = query.Where(x => x.GuaValue == guaValue.Value);
            }

            return await query.ToListAsync();
        }
    }
}