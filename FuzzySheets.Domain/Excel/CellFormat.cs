namespace FuzzySheets.Domain.Excel
{
    public class CellFormat
    {
        public string NumberFormat { get; set; } = string.Empty;
        public string FontName { get; set; } = string.Empty;
        public int FontSize { get; set; }
        public bool Bold { get; set; }
    }
}