using System.Collections.Generic;

namespace FuzzySheets.Domain.Excel
{
    public class Workbook : IWorkbook
    {
        public List<IWorksheet> Worksheets { get; set; } = new List<IWorksheet>();
    }
}