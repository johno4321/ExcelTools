using System;
using System.Linq;
using Microsoft.Office.Interop.Excel;

namespace ExcelRiskTools
{
    public static class ExcelExtensions
    {
        public static void SetCellValue(this Worksheet worksheet, int row, int col, string value)
        {
            var range = worksheet.Cells[row, col] as Range;
            range.set_Value(Type.Missing, value);
        }

        public static bool DoesSheetExist(this Workbook workbook, string worksheetName)
        {
            return (from object worksheet in workbook.Worksheets select worksheet as Worksheet).Any(castedWorksheet => castedWorksheet.Name == worksheetName);
        }

        public static bool AreThereAnySheets(this Workbook workbook)
        {
            if(workbook == null) return false;

            return workbook.Worksheets.Count > 0;
        }

        public static Workbook CreateWorkbook(this Application application)
        {
            var workbook = application.Workbooks.Add(Type.Missing);
            workbook.Activate();
            return workbook;
        }

        public static Worksheet CreateWorksheet(this Workbook workbook)
        {
            var worksheet = workbook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            return worksheet as Worksheet;
        }

        public static Worksheet GetWorksheetByName(this Sheets worksheets, string name)
        {
            return worksheets[name] as Worksheet;
        }
    }
}
