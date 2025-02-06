using System.IO;
using FuzzySheets.Domain.Excel;

namespace FuzzySheets.Services
{
    public interface IParseExcel
    {
        public IWorkbook ParseExcel(Stream file);
    }
}