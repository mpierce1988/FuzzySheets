using System;
using FuzzySheets.Domain.Config;
using FuzzySheets.Domain.Excel;

namespace FuzzySheets.Services.Mutation
{
    public class MutationEngine : IMutationEngine
    {
        public IWorkbook ApplyMutations(IWorkbook excelFile, MutationConfig config)
        {
            throw new NotImplementedException();
        }
    }
}