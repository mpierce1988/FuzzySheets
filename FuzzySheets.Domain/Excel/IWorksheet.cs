using System.Collections.Generic;

namespace FuzzySheets.Domain.Excel
{
    public interface IWorksheet
    {
        public string Name { get; set; }
        public List<List<ICell>> Cells { get; set; }
    }
}