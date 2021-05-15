using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace MicrosoftExcelFileHandler
{
    public class ExcelFileHandler
    {
        // Appends data to the excel document specified by the filename argument
        public static void AppendToExcel(string filename)
        {
            Excel.Application app = new Excel.Application(); // Open an instance of the Microsoft Excel application
            if (app != null) // Run if the app was opened correctly
            {
                // Add a new workbook to the application and open up the first sheet
                Excel.Workbook workbook = app.Workbooks.Add();
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];

                ReadOldData(worksheet, filename); // Grab the old data from the worksheet using the ReadOldData method

                app.ActiveWorkbook.SaveAs(filename, Excel.XlFileFormat.xlWorkbookDefault);

                workbook.Close();
                app.Quit();
            }
        }

        // Reads the old data from the Excel document and adds it to the argument worksheet
        private static void ReadOldData(Excel.Worksheet worksheet, string fileName)
        {
            Excel.Application localApp = new Excel.Application(); // Open an instance of the Microsoft Excel application
            if (localApp != null) // Run if the app was opened correctly
            {
                // Open up a separate worksheet and workbook
                Excel.Workbook oldWorkbook = localApp.Workbooks.Open(fileName);
                Excel.Worksheet oldWorksheet = (Excel.Worksheet)oldWorkbook.Sheets[1];
                
                // Grab information from the old worksheet
                Excel.Range usedRange = oldWorksheet.UsedRange;
                int numRows = usedRange.Rows.Count;
                int numCols = usedRange.Columns.Count;

                worksheet.Name = oldWorksheet.Name; // Set the name of the argument worksheet to the old worksheet name

                // Copy all of the information from the old worksheet to the argument worksheet
                for (int i = 1; i <= numRows; i++)
                {
                    for (int j = 1; j <= numCols; j++)
                    {
                        Excel.Range currRange = (oldWorksheet.Cells[i, j] as Excel.Range);
                        worksheet.Cells[i, j] = currRange.Value == null ? String.Empty : currRange.Value.ToString();
                    }
                }

                oldWorkbook.Close();
                localApp.Quit();
            }
            
        }
    }
}
