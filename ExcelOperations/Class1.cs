using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelHandler;
using NUnit.Framework;

namespace ExcelOperations
{
    public class ExcelTests
    {

        [Test]
        public void SaveExcel()
        {
            var file1 = @"D:\Projects\Testlab\ExcelOperations\ExcelOperations\file.xlsx";
            using (var excelHandler = ExcelHandlerFactory.Instance.Create(file1))
            {
                var sheet = excelHandler.CreateSheet("new sheet");
                Console.WriteLine(sheet.GetCellValue(1, "A"));
                sheet.SetCellValue(1, "A", "Test value in A1 cell");
                Console.WriteLine(sheet.GetCellValue(1, "A"));
                excelHandler.GetSheet(0).SetCellValue(1, "A", 10);
                //excelHandler.g

                excelHandler.Save(@"D:\Projects\Testlab\ExcelOperations\ExcelOperations\file_changed.xlsx");
            }
        }
    }
}
