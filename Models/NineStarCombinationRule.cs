namespace Ecanapi.Models
{
    /// <summary>九星吉凶組合三十三則 KB（主位×勳位的命名規則）</summary>
    public class NineStarCombinationRule
    {
        public int Id { get; set; }
        public int StarA { get; set; }            // 主位星 1-9
        public int StarB { get; set; }            // 勳位星 1-9
        public string Title { get; set; } = "";   // 規則名稱（如「三七蚩尤煞」）
        public string Verdict { get; set; } = ""; // 大吉/吉/偏吉/平/偏凶/凶/大凶
        public string Description { get; set; } = ""; // 克應說明
        public bool? IsProspering { get; set; }   // null=通用, true=僅得運時, false=僅失運時
    }
}
