using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

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

        // This method takes a .xml file (IFormFile) uploaded, read it and populates a voucher list
        public async Task<List<Voucher>> ValidarStatusAfiliadoVoucherAsync(IFormFile file, CancellationToken cancellationToken)
        {
            var vouchers = new List<Voucher>();

            using (var stream = new MemoryStream())
            {
                await file.CopyToAsync(stream, cancellationToken);

                using (var package = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    var rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        var voucher = new Voucher
                        {
                            Name = worksheet.Cells[row, 2]?.Value.ToString(),
                            Company = worksheet.Cells[row, 3]?.Value.ToString()
                        };

                        vouchers.Add(voucher);
                    }
                }
            }

            return vouchers;
        }
    }
}
