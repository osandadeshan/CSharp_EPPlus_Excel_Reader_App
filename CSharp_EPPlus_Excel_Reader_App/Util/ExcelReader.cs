using System.IO;
using OfficeOpenXml;

namespace CSharp_EPPlus_Excel_Reader_App.Util
{
    public class ExcelReader
    {
        private readonly ExcelWorksheet _worksheet;

        public ExcelReader(string excelFilePath, string sheetName)
        {
            var excelFile = new FileInfo(excelFilePath);
            var excelPackage = new ExcelPackage(excelFile);
            _worksheet = excelPackage.Workbook.Worksheets[sheetName];
        }

        public string GetCellValue(int rowNumber, int columnNumber)
        {
            return _worksheet.Cells[rowNumber, columnNumber].Value.ToString();
        }

        public int GetTotalColumnsCount()
        {
            return _worksheet.Dimension.End.Column;
        }

        public int GetTotalRowsCount()
        {
            return _worksheet.Dimension.End.Row;
        }
    }
}