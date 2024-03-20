using Microsoft.AspNetCore.Mvc;
using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;



namespace ExcelProject.Controllers
{
    [Route("api/excel")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        private static int InsertSharedStringItem(string text, SharedStringTablePart sharedStringPart)
        {
            if (sharedStringPart.SharedStringTable == null)
            {
                sharedStringPart.SharedStringTable = new SharedStringTable();
            }

            int index = 0;

            foreach (SharedStringItem item in sharedStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return index;
                }

                index++;
            }

            sharedStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
            sharedStringPart.SharedStringTable.Save();

            return index;
        }


        [HttpPost]
        [Route("edit/{filepath}/{sheetid}/{newValue}/{rowid}/{columnid}")]
        public async Task<IActionResult> EditExcel(String filepath, int sheetid, string newValue, int rowid, int columnid)
        {
            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filepath, true))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart; // whole excel file with all the sheets
                    Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Skip(sheetid).FirstOrDefault(); // any sheet inside the workbook
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id); // Get the SheetData element of the given sheet.Id
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First(); // Get the SheetData element of the given sheet.Id
                    SharedStringTablePart stringTablePart = workbookPart.SharedStringTablePart;
                    Row row = sheetData.Elements<Row>().ElementAt(rowid);
                    Cell cell = row.Elements<Cell>().ElementAt(columnid);
                    cell.CellValue = new CellValue(newValue);
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                    worksheetPart.Worksheet.Save();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

            return Ok();
        }
        
        
        
        [HttpPost]
        [Route("write/{filepath}/{sheetid}")]

        public async Task<IActionResult> WriteExcel(string dataString, string filepath, int sheetid)
        {
            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filepath, true))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart; // whole excel fil    e with all the sheets
                    Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Skip(sheetid).FirstOrDefault(); // any sheet inside the workbook
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id); // Get the SheetData element of the given sheet.Id
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First(); // Get the SheetData element of the given sheet.Id
                    // SharedStringTablePart stringTablePart = spreadsheetDocument.WorkbookPart.SharedStringTablePart;
                    
                    // SharedStringTablePart sharedStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    //
                    // if (sharedStringPart == null)
                    // {
                    //     sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
                    // }
                    //
                    // if (sharedStringPart.SharedStringTable == null)
                    // {
                    //     sharedStringPart.SharedStringTable = new SharedStringTable();
                    // }
                    //     
                    //
                    // string[] data = dataString.Split(",");
                    // Row newRow = new Row();
                    // if (sheetData != null)
                    // {
                    //     foreach(string cellData in data)
                    //     {
                    //         Cell cell = new Cell();
                    //         cell.DataType = CellValues.SharedString;
                    //         int index = InsertSharedStringItem(cellData, sharedStringPart);
                    //         cell.CellValue = new CellValue(index.ToString());
                    //         newRow.AppendChild(cell);   
                    //     }
                    //     sheetData.Append(newRow);
                    //
                    // }
                    // else
                    // {
                    //     Console.WriteLine("No second SheetData element found.");
                    // }
                    
                    
                    string[] data = dataString.Split(",");
                    Row newRow = new Row();
                    if (sheetData != null)
                    {
                        foreach(string cellData in data)
                        {
                            Cell cell = new Cell();
                            cell.DataType = CellValues.String;
                            cell.CellValue = new CellValue(cellData);
                            newRow.AppendChild(cell);
                        }
                        sheetData.Append(newRow);
                    
                    }
                    else
                    {
                        Console.WriteLine("No second SheetData element found.");
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

            return Ok();
        }
        [HttpGet]
        [Route("write/{filepath}/{sheetid}")]
        public async Task<IActionResult> UploadExcel(string filepath, int sheetid)
        {
            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filepath, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart; // whole excel fil    e with all the sheets
                    Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().Skip(sheetid).FirstOrDefault(); // any sheet inside the workbook
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id); // Get the SheetData element of the given sheet.Id
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First(); // Get the SheetData element of the given sheet.Id
                    SharedStringTablePart stringTablePart = workbookPart.SharedStringTablePart;
                    
                    if (sheetData != null)
                    {
                        foreach (Row row in sheetData.Elements<Row>())
                        {
                            foreach (Cell cell in row.Elements<Cell>())
                            {
                                string value = string.Empty;
                                if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                                {
                                    int index = int.Parse(cell.InnerText);
                                    // Console.WriteLine(cell.InnerText);
                                    value = stringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(index).InnerText;
                                }
                                else
                                {
                                    value = cell.CellValue.InnerText;
                                }
                                Console.Write(value + "\t");
                            }
                            Console.WriteLine();
                        }
                    }
                    else
                    {
                        Console.WriteLine("No second SheetData element found.");
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

            return Ok();
        }

    }
}

