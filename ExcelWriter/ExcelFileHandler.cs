using System;
using Excel = Microsoft.Office.Interop.Excel;

/*
 * Resources:
 * https://www.wpf-tutorial.com/dialogs/the-openfiledialog/
 * https://www.delftstack.com/howto/csharp/write-to-excel-in-csharp/
 * https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/interop/how-to-access-office-onterop-objects
 * http://csharpexamples.com/c-read-data-excel-file/
 */
namespace MicrosoftExcelFileHandler
{
    public class ExcelFileHandler
    {
        private const string DATE_FORMAT = "mm/dd/yyyy";
        private const string TIME_FORMAT = "hh:mm AM/PM";
        private const string DECIMAL_FORMAT = "0.00";

        // Appends data to the excel document specified by the filename argument
        public static void AppendToExcel(string filename)
        {
            Excel.Application app = new Excel.Application(); // Open an instance of the Microsoft Excel application
            if (app != null) // Run if the app was opened correctly
            {
                // Add a new workbook to the application and open up the first sheet
                Excel.Workbook workbook = app.Workbooks.Add();
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];

                // Read the old data from the worksheet and write the current date to the first column in the next available row
                int nextAvailableRow = ReadOldData(app, worksheet, filename);
                worksheet.Cells[nextAvailableRow, 1] = DateTime.Now.ToShortDateString();

                // Format the columns based on the data type that they contain
                worksheet.Columns[1].NumberFormat = DATE_FORMAT;
                worksheet.Columns[2].NumberFormat = TIME_FORMAT;
                worksheet.Columns[3].NumberFormat = TIME_FORMAT;
                worksheet.Columns[4].NumberFormat = DECIMAL_FORMAT;
                worksheet.Columns[5].NumberFormat = DECIMAL_FORMAT;

                worksheet.Columns.AutoFit(); // Autofit all of the columns in the worksheet

                try
                {
                    app.ActiveWorkbook.SaveAs(filename, Excel.XlFileFormat.xlWorkbookDefault); // Resave the workbook as the same filename
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    throw new System.Runtime.InteropServices.COMException($"The specified file {filename} could not be saved. Perhaps it is open in another program?");
                }
                finally // The workbook and app will be closed regardless of whether it saved correctly
                {
                    workbook.Close();
                    app.Quit();
                }
            }
        }

        // Reads the old data from the Excel document and adds it to the argument worksheet
        // Returns the number of rows + 1, which indicates the next row to be written
        private static int ReadOldData(Excel.Application app, Excel.Worksheet worksheet, string fileName)
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
            return numRows+1;
        }
    }
}
