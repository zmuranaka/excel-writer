using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace MicrosoftExcelFileHandler
{
    public class ExcelFileHandler
    {
        public static void AppendToExcel(string filename)
        {
            Excel.Application app = new Excel.Application();
            if (app != null)
            {
                Excel.Workbook workbook = app.Workbooks.Add();
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];

                ReadOldData(worksheet, filename);

                app.ActiveWorkbook.SaveAs(filename, Excel.XlFileFormat.xlWorkbookDefault);

                workbook.Close();
                app.Quit();
            }
        }

        private static void ReadOldData(Excel.Worksheet worksheet, string fileName)
        {
            Excel.Application localApp = new Excel.Application();
            if (localApp != null)
            {
                Excel.Workbook oldWorkbook = localApp.Workbooks.Open(fileName);
                Excel.Worksheet oldWorksheet = (Excel.Worksheet)oldWorkbook.Sheets[1];
                Excel.Range usedRange = oldWorksheet.UsedRange;
                int numRows = usedRange.Rows.Count;
                int numCols = usedRange.Columns.Count;
                string oldWorksheetName = oldWorksheet.Name;

                for (int i = 1; i <= numRows; i++)
                {
                    for (int j = 1; j <= numCols; j++)
                    {
                        Excel.Range currRange = (oldWorksheet.Cells[i, j] as Excel.Range);
                        string cellValue = currRange.Value == null ? String.Empty : currRange.Value.ToString();

                        worksheet.Cells[i, j] = cellValue;
                    }
                }

                worksheet.Name = oldWorksheetName;
                oldWorkbook.Close();
            }
            
        }
    }
}
