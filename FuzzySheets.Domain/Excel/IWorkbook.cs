using System.Collections.Generic;

namespace FuzzySheets.Domain.Excel
{
    public interface IWorkbook
    {
        List<IWorksheet> Worksheets { get; set; }
    }
}