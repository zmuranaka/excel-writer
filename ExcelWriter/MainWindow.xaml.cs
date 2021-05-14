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
            fileDialog.Filter = "Excel Files (*.xls;*.xlsx)|*xls;*.xlsx";
            fileDialog.InitialDirectory = "C:\\Temp";
            fileDialog.ShowDialog();
            string fileName = fileDialog.FileName;
            ExcelFileHandler.AppendToExcel(fileName);
        }
    }
}
