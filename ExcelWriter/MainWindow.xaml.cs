﻿using System;
using System.Windows;
using System.Collections.Generic;
using Microsoft.Win32;
using MicrosoftExcelFileHandler;

namespace ExcelWriter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private DateTime startTime;
        private DateTime endTime;
        private List<TimeSpan> sessions = new List<TimeSpan>();

        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Opens up a File Explorer window and saves the data to the file selected.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
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
                    decimal timeTutoring = 0;
                    for (int i = 0; i < sessions.Count; i++)
                    {
                        int totalSecondsInSession = (sessions[i].Hours * 3600) + (sessions[i].Minutes * 60) + sessions[i].Seconds;
                        timeTutoring += totalSecondsInSession / 3600m;
                    }
                    ExcelFileHandler.AppendToExcel(fileName, txtIn.Text, txtOut.Text, decimal.Parse(txtTotal.Text), timeTutoring); // Only call the AppendToExcel method if a valid file was selected
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    MessageBox.Show(ex.Message);
                }
                catch (System.FormatException)
                {
                    MessageBox.Show($"The total \"{txtTotal.Text}\" could not be converted to a decimal. Perhaps it is formatted incorrectly?");
                }
            }
                
        }

        /// <summary>
        /// Remembers the time that the Start button was clicked and displays it in lblTimer
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Start_Click(object sender, RoutedEventArgs e)
        {
            startTime = DateTime.Now;
            lblTimer.Content = startTime.ToString("hh:mm:ss");
        }

        /// <summary>
        /// Remembers the time that the Stop button was clicked.
        /// Displays the difference between the start time and end time in lblTimer.
        /// Adds the difference in times to the sessions List.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Stop_Click(object sender, RoutedEventArgs e)
        {
            endTime = DateTime.Now;
            TimeSpan difference = endTime.Subtract(startTime);
            lblTimer.Content = difference.ToString("hh\\:mm\\:ss");
            sessions.Add(difference); // Add the difference in times to the list of sessions
        }
    }
}
