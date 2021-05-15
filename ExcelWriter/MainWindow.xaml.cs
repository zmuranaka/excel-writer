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
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Excel Files (*.xls;*.xlsx)|*xls;*.xlsx"; // Only allow Microsoft Excel files to be opened
            fileDialog.InitialDirectory = "C:\\Temp"; // The File Explorer window's starting directory is C:\Temp
            fileDialog.ShowDialog();
            string fileName = fileDialog.FileName; // Grab the filename that was opened
            ExcelFileHandler.AppendToExcel(fileName); // Call the AppendToExcel method in the ExcelFileHandler class
        }
    }
}
