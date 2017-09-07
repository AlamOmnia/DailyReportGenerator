using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.IO;
using Demo_Excel_Export.Helpers;
namespace Demo_Excel_Export
{
    class MonthlyIncomOutReport
    {
         private const char DELIMITER = '\t';
        private const string REPORT_TEMPLATE_FILE = "C:/MicroSoft Visual Studio 2015/Demo_Excel_Export/Reports/Fig-2.1_Template.xlsx";
        //private const string REPORT_SHEET_NAME = "Sheet1";
        private const string ADDRESS_OF_IGW_FIRST_CELL = "B13";
        private const string ADDRESS_OF_TOTAL = "A20";
        private const string ADDRESS_OF_DATE = "F8";
        private const int IGW_COUNT = 7;


        private DataTable incomingDataTable, outgoingDataTable;
        private List<IgwModel> igws;
        private string reportDate;

        public MonthlyIncomOutReport(DataTable incomingDataTable, DataTable outgoingDataTable, string  reportDate)
        {
            this.incomingDataTable = incomingDataTable;
            this.outgoingDataTable = outgoingDataTable;
            this.reportDate = "Month:"+reportDate;
            PopulateIgwModels();
        }
     

        public void ExportToExcel(string outputFile)
        {
            // Work on the output copy
            System.IO.File.Copy(REPORT_TEMPLATE_FILE, outputFile, true);

            // Open the spreadsheet document for read-only access.
            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(outputFile, true))
            {
                // Retrieve a reference to the workbook part.
                WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;

                // Find the report sheet name
                IEnumerable<Sheet> sheets = workbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                var sheet = sheets/*.Where(s => s.Name.Equals(REPORT_SHEET_NAME))*/.FirstOrDefault();

                // Retrieve a reference to the worksheet part.
                WorksheetPart worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));


                #region Populate incoming, outgoing tables
                // Use its Worksheet property to get a reference to the cell 
                // whose address matches the address you supplied.
                var cells = worksheetPart.Worksheet.Descendants<Cell>().ToArray<Cell>();
                string cellValue = string.Empty;

                foreach (IgwModel igw in igws)
                {
                    foreach (Cell cell in cells)
                    {
                        cellValue = GetCellValue(spreadSheetDocument, cell);
                        if (IsInIgwColumn(cell.CellReference) && cellValue.Equals(igw.ExcelDisplayName))
                        {
                            InsertIgwValues(igw, cell, worksheetPart);
                        }
                    }
                }
                #endregion

                #region Set totals
                string incomingTotalCallCountFirstCell = GetNextColumnCellReference(ADDRESS_OF_IGW_FIRST_CELL, 1);  // C13
                string incomingTotaLMinsFirstCell = GetNextColumnCellReference(ADDRESS_OF_IGW_FIRST_CELL, 2);  // D13
                string outgoingTotalCallCountFirstCell = GetNextColumnCellReference(ADDRESS_OF_IGW_FIRST_CELL, 3);  // E13
                string outgoingTotaLMinsFirstCell = GetNextColumnCellReference(ADDRESS_OF_IGW_FIRST_CELL, 4);  // F13

                // Totals
                string incomingTotalCallCountSum = GetNextRowCellReference(incomingTotalCallCountFirstCell, IGW_COUNT);  // C20
                string incomingTotalMinsSum = GetNextRowCellReference(incomingTotaLMinsFirstCell, IGW_COUNT);  // D20
                string outgoingTotalCallCountSum = GetNextRowCellReference(outgoingTotalCallCountFirstCell, IGW_COUNT);  // E20
                string outgoingTotalMinsSum = GetNextRowCellReference(outgoingTotaLMinsFirstCell, IGW_COUNT);  // F20


                ExcelHelper.Instance.CalculateSumOfCellRange(workbookPart,
                    worksheetPart,
                    incomingTotalCallCountFirstCell,           // C13
                    GetNextRowCellReference(incomingTotalCallCountFirstCell, IGW_COUNT - 1),    // C19
                    incomingTotalCallCountSum);      // C20 (SUM of c13:c19)

                ExcelHelper.Instance.CalculateSumOfCellRange(workbookPart,
                    worksheetPart,
                    incomingTotaLMinsFirstCell,
                    GetNextRowCellReference(incomingTotaLMinsFirstCell, IGW_COUNT - 1),
                    incomingTotalMinsSum);

                ExcelHelper.Instance.CalculateSumOfCellRange(workbookPart,
                    worksheetPart,
                    outgoingTotalCallCountFirstCell,
                    GetNextRowCellReference(outgoingTotalCallCountFirstCell, IGW_COUNT - 1),
                    outgoingTotalCallCountSum);

                ExcelHelper.Instance.CalculateSumOfCellRange(workbookPart,
                    worksheetPart,
                    outgoingTotaLMinsFirstCell,
                    GetNextRowCellReference(outgoingTotaLMinsFirstCell, IGW_COUNT - 1),
                    outgoingTotalMinsSum);
                #endregion

                #region Set Date
                Cell dateCell = ExcelHelper.Instance.InsertCellInWorksheet(GetColumnName(ADDRESS_OF_DATE), (uint)(GetRowNumber(ADDRESS_OF_DATE)), worksheetPart);

                dateCell.DataType = new EnumValue<CellValues>(CellValues.String);
                dateCell.CellValue = new CellValue(reportDate);
                #endregion
            }
        }

        private void InsertIgwValues(IgwModel igw, Cell igwCell, WorksheetPart worksheetPart)
        {
            uint igwRow = (uint)GetRowNumber(igwCell.CellReference);

            // find incoming & outgoing total Calls & and total mins i.e. C, D, E, F
            string incomingTotalCallsColumn = ((char)((int)GetColumnName(igwCell.CellReference)[0] + 1)).ToString(); // C
            string incomingTotalMinsColumn = ((char)((int)GetColumnName(igwCell.CellReference)[0] + 2)).ToString();   // D
            string outgoingTotalCallsColumn = ((char)((int)GetColumnName(igwCell.CellReference)[0] + 3)).ToString();   // E
            string outgoingTotalMinsColumn = ((char)((int)GetColumnName(igwCell.CellReference)[0] + 4)).ToString();   // F

            // Find the cells to be updated using the column & row number
            Cell incomingTotalCallsCell = ExcelHelper.Instance.InsertCellInWorksheet(incomingTotalCallsColumn, igwRow, worksheetPart);
            Cell incomingTotalMinsCell = ExcelHelper.Instance.InsertCellInWorksheet(incomingTotalMinsColumn, igwRow, worksheetPart);
            Cell outgoingTotalCallsCell = ExcelHelper.Instance.InsertCellInWorksheet(outgoingTotalCallsColumn, igwRow, worksheetPart);
            Cell outgoingTotalMinsCell = ExcelHelper.Instance.InsertCellInWorksheet(outgoingTotalMinsColumn, igwRow, worksheetPart);

            // Set DataTypes to number
            incomingTotalCallsCell.DataType = new EnumValue<CellValues>(CellValues.Number);
            incomingTotalMinsCell.DataType = new EnumValue<CellValues>(CellValues.Number);
            outgoingTotalCallsCell.DataType = new EnumValue<CellValues>(CellValues.Number);
            outgoingTotalMinsCell.DataType = new EnumValue<CellValues>(CellValues.Number);

            // Set values
            incomingTotalCallsCell.CellValue = new CellValue(igw.TotalCallsIncoming);
            incomingTotalMinsCell.CellValue = new CellValue(igw.TotalMinsIncoming);
            //incomingTotalMinsCell.CellFormula = new CellFormula(igw.TotalMinsIncoming);
            outgoingTotalCallsCell.CellValue = new CellValue(igw.TotalCallsOutgoing);
            outgoingTotalMinsCell.CellValue = new CellValue(igw.TotalMinsOutgoing);

            // Save the worksheet.
            worksheetPart.Worksheet.Save();
        }

        private bool IsInIgwColumn(string cellReference)
        {
            string columnName = GetColumnName(cellReference);

            // check the cell is between B13 and B19
            if (columnName[0] == ADDRESS_OF_IGW_FIRST_CELL[0])
            {
                int row = GetRowNumber(cellReference);
                int igwFirstCellRow = GetRowNumber(ADDRESS_OF_IGW_FIRST_CELL);

                if (row >= igwFirstCellRow && row < igwFirstCellRow + IGW_COUNT)
                {
                    return true;
                }
            }

            return false;
        }


        public int GetRowNumber(string cellReference)
        {
            return int.Parse(cellReference.Substring(1));
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

        private string GetNextRowCellReference(string cellReference, uint offset = 1)
        {
            uint nextRowNo = (uint.Parse(cellReference.Substring(1))) + offset;
            return cellReference[0] + nextRowNo.ToString();
        }

        private string GetNextColumnCellReference(string cellReference, uint offset = 1)
        {
            return ((char)((int)GetColumnName(cellReference)[0] + offset)).ToString() + cellReference.Substring(1);
        }

        /*
        private static Cell GetCell(Worksheet worksheet,
             string columnName, uint rowIndex)
        {
            Row row = GetRow(worksheet, rowIndex);

            if (row == null) return null;

            var FirstRow = row.Elements<Cell>().Where(c => string.Compare
                   (c.CellReference.Value, columnName +
                   rowIndex, true) == 0).FirstOrDefault();

            if (FirstRow == null) return null;

            return FirstRow;
        }*/



        private void PopulateIgwModels()
        {
            igws = new List<IgwModel>();
            var sourceNetworks = (
                                     incomingDataTable.AsEnumerable()
                                    .Select(row => row.Field<string>("SourceNetwork"))
                                    .ToArray()
                                 )
                                 .Union
                                 (
                                    outgoingDataTable.AsEnumerable()
                                    .Select(row => row.Field<string>("SourceNetwork"))
                                    .ToArray()
                                 );


            foreach (string sourceNetwork in sourceNetworks)
            {
                IgwModel igw = new IgwModel();
                igw.SourceNetwork = sourceNetwork;

                var incomingIgwDataRow = incomingDataTable.AsEnumerable()
                                  .Where(row => row.Field<string>("SourceNetwork").Equals(sourceNetwork))
                                  .FirstOrDefault();
                var outgoingIgwDataRow = outgoingDataTable.AsEnumerable()
                                  .Where(row => row.Field<string>("SourceNetwork").Equals(sourceNetwork))
                                  .FirstOrDefault();


                if (incomingIgwDataRow != null)
                {
                    igw.TotalCallsIncoming = incomingIgwDataRow["CallCount"] == null ? null : incomingIgwDataRow["CallCount"].ToString();
                    igw.TotalMinsIncoming = incomingIgwDataRow["BilledDuration"] == null ? null : incomingIgwDataRow["BilledDuration"].ToString();
                }

                if (outgoingIgwDataRow != null)
                {
                    igw.TotalCallsOutgoing = outgoingIgwDataRow["CallCount"] == null ? null : outgoingIgwDataRow["CallCount"].ToString();
                    igw.TotalMinsOutgoing = outgoingIgwDataRow["BilledDuration"] == null ? null : outgoingIgwDataRow["BilledDuration"].ToString();
                }

                igw.ExcelDisplayName = GetExcelDisplayName(igw.SourceNetwork);
                igws.Add(igw);
            }
        }


        private string GetExcelDisplayName(string sourceNetwork)
        {
            string line;
            string[] separatedLine;

            using (StringReader reader = new StringReader(Properties.Resources.IgwExcelDisplayToSourceNetworkMapping))
            {
                // Find excel display name for each sourceNetwork name
                while ((line = reader.ReadLine()) != null)
                {
                    separatedLine = line.Split(DELIMITER);

                    // The first column contains the excel name and the 2nd column contains the source network name
                    if (separatedLine.Length == 2)
                    {
                        if (separatedLine[1].Equals(sourceNetwork))
                        {
                            return separatedLine[0];
                        }

                    }
                }
            }

            return string.Empty;
        }
    }
    }

