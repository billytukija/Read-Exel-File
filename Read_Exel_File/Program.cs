using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
using System.Text;

namespace Read_Exel_File
{
    class Program
    {
        static void Main(string[] args)
        {
            string strDoc = @"C:\myfile.xls";

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(strDoc, false))
            {
                WorkbookPart workbookPart = doc.WorkbookPart;
                Sheets thesheetcollection = workbookPart.Workbook.GetFirstChild<Sheets>();
                StringBuilder excelResult = new StringBuilder();

                foreach (Sheet thesheet in thesheetcollection)
                {
                    excelResult.AppendLine("Excel Sheet Name : " + thesheet.Name);
                    excelResult.AppendLine("----------------------------------------------- ");
                    
                    Worksheet theWorksheet = ((WorksheetPart)workbookPart.GetPartById(thesheet.Id)).Worksheet;

                    SheetData thesheetdata = (SheetData)theWorksheet.GetFirstChild<SheetData>();
                    foreach (Row thecurrentrow in thesheetdata)
                    {
                        foreach (Cell thecurrentcell in thecurrentrow)
                        {
                            string currentcellvalue = string.Empty;
                            if (thecurrentcell.DataType != null)
                            {
                                if (thecurrentcell.DataType == CellValues.SharedString)
                                {
                                    int id;
                                    if (Int32.TryParse(thecurrentcell.InnerText, out id))
                                    {
                                        SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                                        if (item.Text != null)
                                        {
                                            excelResult.Append(item.Text.Text + " ");
                                        }
                                        else if (item.InnerText != null)
                                        {
                                            currentcellvalue = item.InnerText;
                                        }
                                        else if (item.InnerXml != null)
                                        {
                                            currentcellvalue = item.InnerXml;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                excelResult.Append(Convert.ToInt16(thecurrentcell.InnerText) + " ");
                            }
                        }
                        excelResult.AppendLine();
                    }
                    excelResult.Append("");
                    Console.WriteLine(excelResult.ToString());
                    Console.ReadLine();
                }
            }
        }
    }
}
