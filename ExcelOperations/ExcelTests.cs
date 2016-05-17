using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelHandler;
using NUnit.Framework;
using OfficeOpenXml;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelOperations
{
    public class ExcelTests
    {
        private const string Drive = "D";
        private const string FileName = Drive + @":\Projects\Testlab\ExcelOperations\ExcelOperations\file.xlsx";

        // adding dropdown in excel for a column
        [Test]
        public void AddExcelDataValidationWithEPPlus()
        {
            if (!File.Exists(FileName))
                throw new InvalidOperationException("No such file!");

            var theFile = new FileInfo(FileName);
            var pck = new ExcelPackage(theFile);
            var excelWorksheets = pck.Workbook.Worksheets;
            var ws = excelWorksheets[1];
            ws.View.ShowGridLines = false;
            
            var values = ws.DataValidations.AddListValidation("A:A");
            values.Formula.Values.Add("1");
            values.Formula.Values.Add("2");
            values.Formula.Values.Add("3");
            values.Formula.Values.Add("4");
            var ms = new MemoryStream();
            pck.SaveAs(ms);
            ms.Position = 0L;

            var fs = new FileStream(Drive + @":\Projects\Testlab\ExcelOperations\ExcelOperations\file_valid.xlsx", FileMode.CreateNew);
            ms.CopyTo(fs);
            fs.Flush();
        }

        [Test]
        public void SimpleChangeWithExcelHandler()
        {
            if (!File.Exists(FileName))
                throw new InvalidOperationException("No such file!");

            using (var excelHandler = ExcelHandlerFactory.Instance.Create(FileName))
            {
                var sheet = excelHandler.CreateSheet("new sheet");
                Console.WriteLine(sheet.GetCellValue(1, "A"));
                sheet.SetCellValue(1, "A", "Test value in A1 cell");
                Console.WriteLine(sheet.GetCellValue(1, "A"));
                excelHandler.GetSheet(1).SetCellValue(2, "G", "Hello World!");
                excelHandler.Save(Drive +  @":\Projects\Testlab\ExcelOperations\ExcelOperations\file_changed.xlsx");
            }
        }
    }
}
