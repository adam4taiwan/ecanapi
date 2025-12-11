namespace Ecanapi.Models
{
    public class CsvDataContainer
    {
        // 用來儲存 data-六十甲子.csv 的查詢結果
        public Dictionary<string, Dictionary<string, string>> LiuShiJiaZiData { get; set; } = new();

        // 用來儲存 data-干支組合.csv 的所有資料 (因為可能需要多組查詢)
        // 這裡建議預先載入整個檔案，因為干支組合的查詢模式是矩陣式的。
        public List<Dictionary<string, string>> GanZhiZuHeAllRows { get; set; }

        // ... 其他 CSV 的資料欄位 ...
    }
}
