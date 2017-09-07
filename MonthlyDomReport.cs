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
       class MonthlyDomReport
        {
            private const char DELIMITER = '\t';
            private const string REPORT_TEMPLATE_FILE = "C:/MicroSoft Visual Studio 2013/Demo_Excel_Export/Reports/Fig-2.2_Template.xlsx";
            //private const string REPORT_SHEET_NAME = "Sheet1";
            private const string ADDRESS_OF_IGW_FIRST_CELL = "B13";
            private const string ADDRESS_OF_TOTAL = "A19";
            private const string ADDRESS_OF_DATE = "F8";
            private const int IGW_COUNT = 6;


            private DataTable domDataTable;
            private List<IgwModel> igws;
            private string reportDate;

            public MonthlyDomReport(DataTable domDataTable, string reportDate)
            {
                this.domDataTable = domDataTable;
                this.reportDate = "Month:" + reportDate;
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
                    string domTotalCallCountFirstCell = GetNextColumnCellReference(ADDRESS_OF_IGW_FIRST_CELL, 1);  // C13
                    string domTotaLMinsFirstCell = GetNextColumnCellReference(ADDRESS_OF_IGW_FIRST_CELL, 2);  // D13

                    // Totals
                    string domTotalCallCountSum = GetNextRowCellReference(domTotalCallCountFirstCell, IGW_COUNT);  // C20
                    string domTotalMinsSum = GetNextRowCellReference(domTotaLMinsFirstCell, IGW_COUNT);  // D20
                   


                    ExcelHelper.Instance.CalculateSumOfCellRange(workbookPart,
                        worksheetPart,
                        domTotalCallCountFirstCell,           // C13
                        GetNextRowCellReference(domTotalCallCountFirstCell, IGW_COUNT - 1),    // C19
                        domTotalCallCountSum);      // C20 (SUM of c13:c19)

                    ExcelHelper.Instance.CalculateSumOfCellRange(workbookPart,
                        worksheetPart,
                        domTotaLMinsFirstCell,
                        GetNextRowCellReference(domTotaLMinsFirstCell, IGW_COUNT - 1),
                        domTotalMinsSum);
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
                string domTotalCallsColumn = ((char)((int)GetColumnName(igwCell.CellReference)[0] + 1)).ToString(); // C
                string domTotalMinsColumn = ((char)((int)GetColumnName(igwCell.CellReference)[0] + 2)).ToString();   // D
                  // F

                // Find the cells to be updated using the column & row number
                Cell domTotalCallsCell = ExcelHelper.Instance.InsertCellInWorksheet(domTotalCallsColumn, igwRow, worksheetPart);
                Cell domTotalMinsCell = ExcelHelper.Instance.InsertCellInWorksheet(domTotalMinsColumn, igwRow, worksheetPart);
                

                // Set DataTypes to number
                domTotalCallsCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                domTotalMinsCell.DataType = new EnumValue<CellValues>(CellValues.Number);
              

                // Set values
                domTotalCallsCell.CellValue = new CellValue(igw.TotalCallsIncoming);
                domTotalMinsCell.CellValue = new CellValue(igw.TotalMinsIncoming);
                //incomingTotalMinsCell.CellFormula = new CellFormula(igw.TotalMinsIncoming);
              

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


            private void PopulateIgwModels()
            {
                igws = new List<IgwModel>();
                var sourceNetworks = (
                                         domDataTable.AsEnumerable()
                                        .Select(row => row.Field<string>("SourceNetwork"))
                                        .ToArray()
                                     );


                foreach (string sourceNetwork in sourceNetworks)
                {
                    IgwModel igw = new IgwModel();
                    igw.SourceNetwork = sourceNetwork;

                    var DomIgwDataRow = domDataTable.AsEnumerable()
                                      .Where(row => row.Field<string>("SourceNetwork").Equals(sourceNetwork))
                                      .FirstOrDefault();
                    


                    if (DomIgwDataRow != null)
                    {
                        igw.TotalCallsIncoming = DomIgwDataRow["CallCount"] == null ? null : DomIgwDataRow["CallCount"].ToString();
                        igw.TotalMinsIncoming = DomIgwDataRow["BilledDuration"] == null ? null : DomIgwDataRow["BilledDuration"].ToString();
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

