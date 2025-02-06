using System;
using System.Collections.Generic;
using System.Linq;
using FuzzySheets.Domain.Config;
using FuzzySheets.Domain.Excel;
using FuzzySheets.Services.Mutation.Rules;

namespace FuzzySheets.Services.Mutation
{
    public class MutationEngine : IMutationEngine
    {
        #region Private Fields

        private readonly Random _random = new Random();

        #endregion

        #region Public Methods

        /// <summary>
        /// Applies mutations to an Excel workbook based on the provided configuration.
        /// </summary>
        /// <param name="excelFile">The Excel workbook to mutate.</param>
        /// <param name="config">The configuration specifying which mutations to apply.</param>
        /// <returns>The mutated Excel workbook.</returns>
        public IWorkbook ApplyMutations(IWorkbook excelFile, MutationConfig config)
        {
            // Validate ExcelFile contains sheets
            if (excelFile.Worksheets.Count <= 0)
            {
                return excelFile;
            }

            Dictionary<MutationDetail, IMutationRule> mutations = GetEnabledMutations(config);

            if (mutations.Count <= 0)
            {
                return excelFile;
            }

            // Loop through worksheets
            foreach (var sheet in excelFile.Worksheets)
            {
                // If the worksheet contains no cells, skip it
                if (sheet.Cells.Count <= 0) continue;

                // Loop through each mutation
                foreach (var mutationDetailRule in mutations)
                {
                    MutationDetail mutationDetail = mutationDetailRule.Key;
                    IMutationRule mutationRule = mutationDetailRule.Value;

                    var cellsToMutate = GetCellsToMutate(sheet, mutationDetail);

                    ApplyMutation(cellsToMutate, mutationRule);
                }
            }

            return excelFile;
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Retrieves the enabled mutations from the provided configuration.
        /// </summary>
        /// <param name="config">The configuration specifying which mutations to apply.</param>
        /// <returns>A dictionary where the key is the mutation detail and the value is the corresponding mutation rule.</returns>
        private Dictionary<MutationDetail, IMutationRule> GetEnabledMutations(MutationConfig config)
        {
            Dictionary<MutationDetail, IMutationRule> mutations = new Dictionary<MutationDetail, IMutationRule>();

            if (config.Encoding.IsEnabled)
            {
                mutations.Add(config.Encoding, new TextEncodingMutation());
            }

            if (config.DateFormat.IsEnabled)
            {
                mutations.Add(config.DateFormat, new DateFormatMutation());
            }

            if (config.NumberFormat.IsEnabled)
            {
                mutations.Add(config.NumberFormat, new NumberFormatMutation());
            }

            if (config.DecimalPrecision.IsEnabled)
            {
                mutations.Add(config.DecimalPrecision, new DecimalPrecisionMutation());
            }

            if (config.EmptyValues.IsEnabled)
            {
                mutations.Add(config.EmptyValues, new EmptyValuesMutation());
            }

            return mutations;
        }

        /// <summary>
        /// Applies the specified mutation rule to a list of cells.
        /// </summary>
        /// <param name="cellsToMutate">The list of cells to mutate.</param>
        /// <param name="mutationRule">The mutation rule to apply to each cell.</param>
        private void ApplyMutation(List<ICell> cellsToMutate, IMutationRule mutationRule)
        {
            foreach (var cell in cellsToMutate)
            {
                var mutatedCell = mutationRule.Mutate(cell);
                cell.Value = mutatedCell.Value;
                cell.Format = mutatedCell.Format;
                cell.Formula = mutatedCell.Formula;
            }
        }

        #endregion

        #region Get Cells Methods

        /// <summary>
        /// Collects cells to mutate from a worksheet based on the mutation details provided.
        /// </summary>
        /// <param name="sheet">The worksheet to collect cells from.</param>
        /// <param name="mutationDetail">The details of the mutation to apply.</param>
        /// <returns>A list of cells to mutate.</returns>
        private List<ICell> GetCellsToMutate(IWorksheet sheet, MutationDetail mutationDetail)
        {
            // Collect cells to mutate
            List<ICell> cellsToMutate = new List<ICell>();

            if (mutationDetail.IsNumeric)
            {
                // Get all cells that are numeric
                List<ICell> numericCells = GetNumericCells(sheet.Cells);

                cellsToMutate.AddRange(GetRandomCells(numericCells, mutationDetail.PercentMutated));
            }

            if (mutationDetail.IsDate)
            {
                // Get all cells that are dates
                List<ICell> dateCells = GetDateCells(sheet.Cells);

                cellsToMutate.AddRange(GetRandomCells(dateCells, mutationDetail.PercentMutated));
            }

            if (mutationDetail.IsString)
            {
                // Get all cells that are strings
                List<ICell> stringCells = GetStringCells(sheet.Cells);

                cellsToMutate.AddRange(GetRandomCells(stringCells, mutationDetail.PercentMutated));
            }

            return cellsToMutate;
        }

        /// <summary>
        /// Collects all cells that contain string values from the provided worksheet cells.
        /// </summary>
        /// <param name="sheetCells">A list of rows, where each row is a list of cells in the worksheet.</param>
        /// <returns>A list of cells that contain string values.</returns>
        private List<ICell> GetStringCells(List<List<ICell>> sheetCells)
        {
            List<ICell> stringCells = new List<ICell>();

            // Loop through each row
            foreach (var row in sheetCells)
            {
                // Loop through each cell in the row
                foreach (var cell in row)
                {
                    if (cell.Value is string)
                    {
                        stringCells.Add(cell);
                    }
                }
            }

            return stringCells;
        }

        /// <summary>
        /// Collects all cells that contain date values from the provided worksheet cells.
        /// </summary>
        /// <param name="sheetCells">A list of rows, where each row is a list of cells in the worksheet.</param>
        /// <returns>A list of cells that contain date values.</returns>
        private List<ICell> GetDateCells(List<List<ICell>> sheetCells)
        {
            List<ICell> dateCells = new List<ICell>();

            // Loop through each row
            foreach (var row in sheetCells)
            {
                // Loop through each cell in the row
                foreach (var cell in row)
                {
                    if (cell.Value is DateTime)
                    {
                        dateCells.Add(cell);
                    }
                }
            }

            return dateCells;
        }

        /// <summary>
        /// Collects all cells that contain numeric values from the provided worksheet cells.
        /// </summary>
        /// <param name="sheetCells">A list of rows, where each row is a list of cells in the worksheet.</param>
        /// <returns>A list of cells that contain numeric values.</returns>
        private List<ICell> GetNumericCells(List<List<ICell>> sheetCells)
        {
            List<ICell> numericCells = new List<ICell>();

            // Loop through each row
            foreach (var row in sheetCells)
            {
                // Loop through each cell in the row
                foreach (var cell in row)
                {
                    if (cell.Value is double || cell.Value is int || cell.Value is decimal)
                    {
                        numericCells.Add(cell);
                    }
                }
            }

            return numericCells;
        }

        /// <summary>
        /// Selects a random subset of cells from the provided list based on the specified percentage to mutate.
        /// </summary>
        /// <param name="cells">The list of cells to select from.</param>
        /// <param name="percentMutated">The percentage of cells to mutate.</param>
        /// <returns>A list of randomly selected cells to mutate.</returns>
        private List<ICell> GetRandomCells(List<ICell> cells, double percentMutated)
        {
            // Randomly shuffle the cells
            cells = cells.OrderBy(x => _random.Next()).ToList();

            // Calculate the number of cells to mutate
            int numCellsToMutate = (int)Math.Ceiling(cells.Count * percentMutated);

            return cells.Take(numCellsToMutate).ToList();
        }

        #endregion
    }
}