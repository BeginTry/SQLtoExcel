using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Packaging;

namespace SQLtoExcel
{
    class Utils
    {
        /// <summary>
        /// Creates a Microsoft Excel Worksheet from the input DataSet, returning it as a MemoryStream.
        /// </summary>
        /// <param name="ds">DataSet of source data.</param>
        /// <returns>MemoryStream</returns>
        public static MemoryStream ExportDataSetToExcel(DataSet ds)
        {
            //Initial code source:
            //https://accesspublic.wordpress.com/2014/02/22/c-export-dataset-to-excel-using-openxml/

            MemoryStream stream = new MemoryStream();
            using (var workbook = SpreadsheetDocument.Create(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = workbook.AddWorkbookPart();
                workbook.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
                workbook.WorkbookPart.Workbook.Sheets = new DocumentFormat.OpenXml.Spreadsheet.Sheets();

                for (int tableIdx = 0; tableIdx < ds.Tables.Count; tableIdx++)
                {
                    DataTable table = ds.Tables[tableIdx];
                    var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new DocumentFormat.OpenXml.Spreadsheet.SheetData();
                    sheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                    DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbook.WorkbookPart.Workbook.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Sheets>();
                    string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);
                    DocumentFormat.OpenXml.Spreadsheet.Sheet sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet();

                    sheet.Id = relationshipId;
                    sheet.SheetId = (uint)(tableIdx + 1);   //If set to zero, Excel will display an error when opening the spreadsheet file.
                    sheet.Name = table.TableName;
                    sheets.Append(sheet);

                    DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                    List<String> columns = new List<string>();

                    foreach (System.Data.DataColumn column in table.Columns)
                    {
                        columns.Add(column.ColumnName);
                        DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();
                        cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                        cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(column.ColumnName);
                        headerRow.AppendChild(cell);
                    }

                    
                    //DocumentFormat.OpenXml.Spreadsheet.Pane p = new DocumentFormat.OpenXml.Spreadsheet.Pane()
                    //{
                    //    VerticalSplit = 9D,
                    //    TopLeftCell = "A2",
                    //    ActivePane = DocumentFormat.OpenXml.Spreadsheet.PaneValues.BottomLeft,
                    //    State = DocumentFormat.OpenXml.Spreadsheet.PaneStateValues.Frozen
                    //};


                    //sheetData.AppendChild(p);

                    sheetData.AppendChild(headerRow);


                    foreach (System.Data.DataRow dsrow in table.Rows)
                    {
                        DocumentFormat.OpenXml.Spreadsheet.Row newRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                        foreach (String col in columns)
                        {
                            DocumentFormat.OpenXml.Spreadsheet.Cell cell = new DocumentFormat.OpenXml.Spreadsheet.Cell();

                            switch (Type.GetTypeCode(table.Columns[col].DataType))
                            {
                                case System.TypeCode.Int16:
                                case System.TypeCode.Int32:
                                case System.TypeCode.Int64:
                                case System.TypeCode.UInt16:
                                case System.TypeCode.UInt32:
                                case System.TypeCode.UInt64:
                                    cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
                                    break;
                                case System.TypeCode.Decimal:
                                case System.TypeCode.Double:
                                case System.TypeCode.Single:
                                    cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Number;
                                    break;
                                case System.TypeCode.DateTime:
                                    cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.Date;
                                    break;
                                default:
                                    cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                                    break;
                            }


                            //cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
                            cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dsrow[col].ToString()); //
                            newRow.AppendChild(cell);
                        }

                        sheetData.AppendChild(newRow);
                    }

                }
            }
            return stream;
        }
    }
}
