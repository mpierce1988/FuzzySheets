namespace FuzzySheets.Domain.Excel
{
    public interface ICell
    {
        public object? Value { get; set; }
        public string Formula { get; set; }
        public CellFormat Format { get; set; }
    }
}