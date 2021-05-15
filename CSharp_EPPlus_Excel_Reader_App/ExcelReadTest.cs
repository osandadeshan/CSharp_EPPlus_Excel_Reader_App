using System;
using System.Reflection;
using CSharp_EPPlus_Excel_Reader_App.Util;
using NUnit.Framework;

namespace CSharp_EPPlus_Excel_Reader_App
{
    [TestFixture]
    public class ExcelReadTest
    {
        private string _excelFilePath;
        
        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            var path = Assembly.GetCallingAssembly().CodeBase;
            if (path == null) return;
            var projectPath = new Uri(path.Substring(0,
                path.LastIndexOf("bin", StringComparison.Ordinal))).LocalPath;
            _excelFilePath = new Uri(projectPath).LocalPath + @"PersonDetails.xlsx";
        }

        [Test]
        public void TestPersonInfoSheet()
        {
            var personInfoSheetReader = new ExcelReader(_excelFilePath, "Person_Info");

            var b4CellValue = personInfoSheetReader.ReadCellValue(4, 2);
            Console.WriteLine("B4 (4th Row, 2nd Column) Cell Value: " + b4CellValue);
            Assert.AreEqual("Gemunu", b4CellValue);

            var d3CellValue = personInfoSheetReader.ReadCellValue(3, 4);
            Console.WriteLine("D3 (3rd Row, 4th Column) Cell Value: " + d3CellValue);
            Assert.AreEqual("Singapore", d3CellValue);
        }

        [Test]
        public void TestContactInfoSheet()
        {
            var contactInfoSheetReader = new ExcelReader(_excelFilePath, "Contact_Info");

            var c4CellValue = contactInfoSheetReader.ReadCellValue(4, 3);
            Console.WriteLine("C4 (4th Row, 3rd Column) Cell Value: " + c4CellValue);
            Assert.AreEqual("Thissamaharama", c4CellValue);

            var b3CellValue = contactInfoSheetReader.ReadCellValue(3, 2);
            Console.WriteLine("B3 (3rd Row, 2nd Column) Cell Value: " + b3CellValue);
            Assert.AreEqual("714567891", b3CellValue);
        }

        [Test]
        public void TestPrintAllData()
        {
            var personInfoSheetReader = new ExcelReader(_excelFilePath, "Person_Info");
            
            for (var row = 1; row <= personInfoSheetReader.GetTotalRowsCount(); row++)
            {
                for (var col = 1; col <= personInfoSheetReader.GetTotalColumnsCount(); col++)
                {
                    Console.WriteLine(" Row:" + row + " Column:" + col + " Value:" +
                                      personInfoSheetReader.ReadCellValue(row, col)?.Trim());
                }
            }
        }
    }
}