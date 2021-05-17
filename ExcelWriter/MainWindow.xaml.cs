using System;
using System.Windows;
using Microsoft.Win32;
using MicrosoftExcelFileHandler;

namespace ExcelWriter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        DateTime startTime;
        DateTime endTime;
        TimeSpan difference;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Excel Files (*.xls;*.xlsx)|*xls;*.xlsx"; // Only allow Microsoft Excel files to be opened
            fileDialog.InitialDirectory = "C:\\Temp"; // The File Explorer window's starting directory is C:\Temp
            fileDialog.ShowDialog();
            string fileName = fileDialog.FileName; // Grab the filename that was opened
            if (!string.IsNullOrEmpty(fileName))
            {
                try
                {
                    ExcelFileHandler.AppendToExcel(fileName); // Only call the AppendToExcel method if a valid file was selected
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
                
        }

        private void Start_Click(object sender, RoutedEventArgs e)
        {
            lblTimer.Visibility = Visibility.Visible;
            startTime = DateTime.Now;
            lblTimer.Content = startTime.ToString("hh:mm:ss");
        }

        private void Stop_Click(object sender, RoutedEventArgs e)
        {
            endTime = DateTime.Now;
            difference = endTime.Subtract(startTime);
            lblTimer.Content = difference.ToString("hh\\:mm\\:ss");
        }
    }
}
