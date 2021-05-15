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

                ReadOldData(app, worksheet, filename); // Grab the old data from the worksheet using the ReadOldData method

                app.ActiveWorkbook.SaveAs(filename, Excel.XlFileFormat.xlWorkbookDefault); // Resave the workbook as the same filename

                workbook.Close();
                app.Quit();
            }
        }

        // Reads the old data from the Excel document and adds it to the argument worksheet
        private static void ReadOldData(Excel.Application app, Excel.Worksheet worksheet, string fileName)
        {
            // Open up a separate worksheet and workbook
            Excel.Workbook oldWorkbook = app.Workbooks.Open(fileName);
            Excel.Worksheet oldWorksheet = (Excel.Worksheet)oldWorkbook.Sheets[1];
                
            // Grab information from the old worksheet
            int numRows = oldWorksheet.UsedRange.Rows.Count;
            int numCols = oldWorksheet.UsedRange.Columns.Count;

            worksheet.Name = oldWorksheet.Name; // Set the name of the argument worksheet to the old worksheet name

            // Copy all of the information from the old worksheet to the argument worksheet
            for (int row = 1; row <= numRows; row++)
            {
                for (int col = 1; col <= numCols; col++)
                {
                    Excel.Range currRange = (oldWorksheet.Cells[row, col] as Excel.Range);
                    worksheet.Cells[row, col] = currRange.Value == null ? String.Empty : currRange.Value.ToString();
                }
            }

            oldWorkbook.Close();
        }
    }
}
