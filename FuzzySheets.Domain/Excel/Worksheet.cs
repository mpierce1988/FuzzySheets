using System.Collections.Generic;

namespace FuzzySheets.Domain.Excel
{
    public class Worksheet : IWorksheet
    {
        public string Name { get; set; } = string.Empty;
        public List<List<ICell>> Cells { get; set; } = new List<List<ICell>>();
    }
}