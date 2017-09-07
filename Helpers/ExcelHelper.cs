using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace Demo_Excel_Export.Helpers
{
    /// <summary>
    /// Helper for parsing and loading excel files (.xlsx) to database
    /// </summary>
    class ExcelHelper
    {
        #region Singleton implementation
        private static ExcelHelper excelHelper;
        public static ExcelHelper Instance
        {
            get
            {
                if (excelHelper == null)
                    excelHelper = new ExcelHelper();
                return excelHelper;
            }
        }
        #endregion

        /// <summary>
        /// Cell value retrival method from MSDN website.
        /// </summary>
        /// <param name="document">The excel document that is being parsed.</param>
        /// <param name="cell">The cell from which value is being extracted.</param>
        /// <returns>Returns a string that contains the value of the given cell.</returns>
        private string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            string value = string.Empty;
            // If the cell does not exist, return an empty string.
            if (cell != null)
            {
                CellValue cellValue = cell.CellValue;
                value = (cellValue == null) ? cell.InnerText : cellValue.Text;

                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    var stringTable = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    if (stringTable != null)
                        value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                }
            }
            return value;
        }

        // Given a document name, a worksheet name, the name of the first cell in the contiguous range, 
        // the name of the last cell in the contiguous range, and the name of the results cell, 
        // calculates the sum of the cells in the contiguous range and inserts the result into the results cell.
        // Note: All cells in the contiguous range must contain numbers.
        public void CalculateSumOfCellRange(WorkbookPart workbookPart, WorksheetPart worksheetPart, string firstCellRefrenece, string lastCellReference, string resultCellReference, bool useFormulaValue = false)
        {
            Worksheet worksheet = worksheetPart.Worksheet;

            // Get the row number and column name for the first and last cells in the range.
            uint firstRowNum = GetRowNumber(firstCellRefrenece);
            uint lastRowNum = GetRowNumber(lastCellReference);
            string firstColumn = GetColumnName(firstCellRefrenece);
            string lastColumn = GetColumnName(lastCellReference);

            double sum = 0;

            // Iterate through the cells within the range and add their values to the sum.
            foreach (Row row in worksheet.Descendants<Row>().Where(r => r.RowIndex.Value >= firstRowNum && r.RowIndex.Value <= lastRowNum))
            {
                foreach (Cell cell in row)
                {
                    string columnName = GetColumnName(cell.CellReference.Value);
                    if (CompareColumn(columnName, firstColumn) >= 0 && CompareColumn(columnName, lastColumn) <= 0)
                    {
                        if (useFormulaValue)
                        {
                            var formula = cell.CellFormula.ToString();
                            if (!string.IsNullOrEmpty(formula))
                            {
                                sum += double.Parse(formula);
                            }
                        }
                        else
                        {
                            var value = GetCellValue(workbookPart, cell);
                            if (!string.IsNullOrEmpty(value))
                            {
                                sum += double.Parse(value);
                            }
                        }
                    }
                }
            }

            Cell result = InsertCellInWorksheet(GetColumnName(resultCellReference), GetRowNumber(resultCellReference), worksheetPart);

            // Set the value of the cell.
            result.DataType = new EnumValue<CellValues>(CellValues.Number);
            result.CellValue = new CellValue(sum.ToString());

            worksheetPart.Worksheet.Save();
        }

        private string GetCellValue(WorkbookPart workbookPart, Cell cell)
        {
            string value = string.Empty;
            // If the cell does not exist, return an empty string.
            if (cell != null)
            {
                CellValue cellValue = cell.CellValue;
                value = (cellValue == null) ? cell.InnerText : cellValue.Text;

                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    if (stringTable != null)
                        value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                }
            }
            return value;
        }

       
        // Given two columns, compares the columns.
        private int CompareColumn(string column1, string column2)
        {
            if (column1.Length > column2.Length)
            {
                return 1;
            }
            else if (column1.Length < column2.Length)
            {
                return -1;
            }
            else
            {
                return string.Compare(column1, column2, true);
            }
        }

        public uint GetRowNumber(string cellReference)
        {
            return uint.Parse(cellReference.Substring(1));
        }

        /// <summary>
        /// Given a cell name, parses the specified cell to get the column name.
        /// </summary>
        /// <param name="cellReference">Address of the cell (ie. B2)</param>
        /// <returns>Column Name (ie. B)</returns>
        private string GetColumnName(string cellReference)
        {
            StringBuilder stringBuilder = new StringBuilder();
            int i = 0;

            while (!char.IsDigit(cellReference[i]))
            {
                stringBuilder.Append(cellReference[i++]);
            }

            return stringBuilder.ToString();
        }

        /// Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        // If the cell already exists, returns it. 
        public Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }

        public void ExportToExcel(DataSet ds, string file)
        {
            using (var workbook = SpreadsheetDocument.Create(file, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = workbook.AddWorkbookPart();
                workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
                workbook.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();

                uint sheetId = 1;

                foreach (DataTable table in ds.Tables)
                {
                    var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
                    sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                    DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
                    string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                    if (sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Count() > 0)
                    {
                        sheetId =
                            sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                    }

                    DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet() { Id = relationshipId, SheetId = sheetId, Name = table.TableName };
                    sheets.Append(sheet);

                    DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();

                    List<string> columns = new List<string>();
                    foreach (DataColumn column in table.Columns)
                    {
                        columns.Add(column.ColumnName);

                        DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                        cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                        cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName);
                        headerRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(headerRow);

                    foreach (DataRow dsrow in table.Rows)
                    {
                        DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                        foreach (string col in columns)
                        {
                            DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                            cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString()); //
                            newRow.AppendChild(cell);
                        }

                        sheetData.AppendChild(newRow);
                    }
                }
            }
        }
    }
}
