using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace TotalCalculator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string phonesTextPath = @"D:\Dropbox\Text Files\Phones.txt";
        string phonesExcelPath = @"D:\Dropbox\Text Files\Phones.xlsx";
        string savingPath = @"D:\Dropbox\Text Files\iphone.txt";
        string dolaratPath = @"D:\Dropbox\Grandstream new\Progs\JaroorDolarat.txt";
        string jaroor1Path = @"D:\Dropbox\Text Files\globe7 accounts.txt";

        double touchDifference = 1 - (1105.0 / 1500);
        double alfaDifference = 1 - (1115.0 / 1500);

        int losses = 600;  // in $

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            double totalDolarat = 0;
            double touchDollars = 0, alfaDollars = 0;
            double phone1 = 0, phone2 = 0;
            double savings = 0;
            double jaroor1 = 0;

            Task t = Task.Run(() =>
            {

                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(phonesExcelPath);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                phone2 = xlRange.Cells[25, 4].Value2;

                this.Dispatcher.Invoke(() =>
                {
                    totalDolarat += phone2;
                    label.Content = totalDolarat.ToString("C2");
                    status.Content = "Done";
                });

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);

            });

            string Data = File.ReadAllText(dolaratPath);
            MatchCollection matches;
            matches = Regex.Matches(Data, @"total   = (\d{1,4}(?:\.\d{1,2})?)", RegexOptions.Singleline);
            totalDolarat = double.Parse(matches[matches.Count - 1].Groups[1].Value);
            matches = Regex.Matches(Data, @"(\d{1,4}(?:\.\d{1,2})?)\(MTC\) \+ (\d{1,4}(?:\.\d{1,2})?)\(Alfa\)", RegexOptions.Singleline);
            touchDollars = double.Parse(matches[matches.Count - 1].Groups[1].Value);
            alfaDollars = double.Parse(matches[matches.Count - 1].Groups[2].Value);
            totalDolarat -= losses + (touchDollars * touchDifference) + (alfaDollars * alfaDifference);

            Data = File.ReadAllText(phonesTextPath);
            Match mat = Regex.Match(Data, @"\= (\d{3,5}\.?5?)");
            phone1 = double.Parse(mat.Groups[1].Value);
            phone1 /= 1.5;

            Data = File.ReadAllText(savingPath);
            mat = Regex.Match(Data, @" (\d{2,4})\+(\d{2,5})\+(\d{2,4})\$");
            savings += double.Parse(mat.Groups[1].Value);
            savings += double.Parse(mat.Groups[2].Value);
            savings += double.Parse(mat.Groups[3].Value);

            Data = File.ReadAllLines(jaroor1Path).Last();
            mat = Regex.Match(Data, @"\= (\d{2,4})");
            jaroor1 = double.Parse(mat.Groups[1].Value);
            jaroor1 += 70;
            jaroor1 /= 1.5;


            totalDolarat = totalDolarat + phone1 + savings + jaroor1;

            label.Content = totalDolarat.ToString("C2");

        }
    }
}
