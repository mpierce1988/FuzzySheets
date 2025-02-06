namespace FuzzySheets.Domain.Excel
{
    public class Cell : ICell
    {
        public object? Value { get; set; }
        public string Formula { get; set; } = string.Empty;
        public CellFormat Format { get; set; } = new CellFormat();
    }
}