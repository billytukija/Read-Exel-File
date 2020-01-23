using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using System;
using System.IO;
using System.Linq;
using System.Text;

namespace Read_Exel_File
{
    class Program
    {
        static void Main(string[] args)
        {
            string strDoc = @"C:\1579630636063342.xls";

            if (!strDoc.EndsWith(".xls"))
                throw new Exception("Importa um arquivo xml");

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = File.Open(strDoc, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                        var counter = 0;
                        while (reader.Read()) //Each ROW
                        {
                            for (int column = 0; column < reader.FieldCount; column++)
                            {
                                //Console.WriteLine(reader.GetString(column));//Will blow up if the value is decimal etc. 
                                Console.WriteLine(reader.GetValue(column));//Get Value returns object

                            }

                            counter++;

                            Console.WriteLine("Line " + counter);

                        }
                    } while (reader.NextResult()); //Move to NEXT SHEET

                }
            }
        }
    }
}
