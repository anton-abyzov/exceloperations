using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelHandler;
using NUnit.Framework;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelOperations
{
    public class ExcelTests
    {

        [Test]
        public void SaveExcel()
        {
            var file1 = @"C:\Projects\Testlab\ExcelOperations\ExcelOperations\file.xlsx";
            using (var excelHandler = ExcelHandlerFactory.Instance.Create(file1))
            {
                var sheet = excelHandler.CreateSheet("new sheet");
                Console.WriteLine(sheet.GetCellValue(1, "A"));
                sheet.SetCellValue(1, "A", "Test value in A1 cell");
                Console.WriteLine(sheet.GetCellValue(1, "A"));
                excelHandler.GetSheet(1).SetCellValue(2, "G", "df, sdfsf, dd");
                //excelHandler.GetSheet(1).

                //Excel.Application t = new Excel.ApplicationClass();

                excelHandler.Save(@"C:\Projects\Testlab\ExcelOperations\ExcelOperations\file_changed.xlsx");
            }
        }
    }
}
