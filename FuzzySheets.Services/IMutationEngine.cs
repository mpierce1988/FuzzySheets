using FuzzySheets.Domain.Config;
using FuzzySheets.Domain.Excel;

namespace FuzzySheets.Services
{
    public interface IMutationEngine
    {
        public IWorkbook ApplyMutations(IWorkbook excelFile, MutationConfig config);
    }
}