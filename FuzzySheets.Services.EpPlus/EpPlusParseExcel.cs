using System;
using System.Collections.Generic;
using System.IO;
using FuzzySheets.Domain.Excel;
using OfficeOpenXml;

namespace FuzzySheets.Services.EpPlus
{
    public class EpPlusParseExcel : IParseExcel
    {
        #region Public Methods

        /// <summary>
        /// Parses an Excel file stream and extracts its data into a Workbook object.
        /// </summary>
        /// <param name="file">The stream of the Excel file to parse.</param>
        /// <returns>A <see cref="IWorkbook"/> object containing the parsed workbook data.</returns>
        public IWorkbook ParseExcel(Stream file)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var workbook = new Workbook();

            var package = new ExcelPackage(file);

            // Loop through each worksheet
            foreach (var sheet in package.Workbook.Worksheets)
            {
                Worksheet worksheet = ParseWorksheet(sheet);

                workbook.Worksheets.Add(worksheet);
            }

            return workbook;
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Parses an Excel worksheet and extracts its data into a Worksheet object.
        /// </summary>
        /// <param name="sheet">The Excel worksheet to parse.</param>
        /// <returns>A <see cref="Worksheet"/> object containing the parsed worksheet data.</returns>
        private Worksheet ParseWorksheet(ExcelWorksheet sheet)
        {
            Worksheet worksheet = new Worksheet()
            {
                Name = sheet.Name
            };

            // Ensure the sheet has data. If no data, continue to the next sheet
            if (sheet.Dimension == null) return worksheet;

            // Loop through each row
            for (int row = sheet.Dimension.Start.Row; row <= sheet.Dimension.End.Row; row++)
            {
                List<ICell> parsedRow = new List<ICell>();

                // Loop through each column
                for (int col = sheet.Dimension.Start.Column; col <= sheet.Dimension.End.Column; col++)
                {
                    ExcelRange excelCell = sheet.Cells[row, col];

                    parsedRow.Add(ParseCell(excelCell));
                }

                worksheet.Cells.Add(parsedRow);
            }

            return worksheet;
        }

        /// <summary>
        /// Parses an Excel cell and extracts its value, formula, and formatting.
        /// </summary>
        /// <param name="cell">The Excel cell to parse.</param>
        /// <returns>A <see cref="Cell"/> object containing the parsed cell data.</returns>
        private Cell ParseCell(ExcelRange cell)
        {
            // Extract cell formatting
            CellFormat cellFormat = new CellFormat()
            {
                Bold = cell.Style.Font.Bold,
                FontName = cell.Style.Font.Name,
                FontSize = Convert.ToInt32(cell.Style.Font.Size),
                NumberFormat = cell.Style.Numberformat.Format
            };

            // Create a new Cell object with the extracted value, formula, and format
            Cell parsedCell = new Cell()
            {
                Value = cell.Value,
                Formula = cell.Formula,
                Format = cellFormat
            };

            return parsedCell;
        }

        #endregion
    }
}