using FuzzySheets.Domain.Excel;

namespace FuzzySheets.Services
{
    public interface IMutationRule
    {
        public ICell Mutate(ICell cell);
    }
}