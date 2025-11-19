using EWPFDesktop;
using EWPFDesktop.RPC;
using Microsoft.Win32;
using RabbitMQ.Client;
using System;
using System.IO;
using System.Net;
using System.Threading;
using System.Windows;
using System.Windows.Shapes;
using static System.Net.WebRequestMethods;

namespace ExcelToWord
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        bool availabilityTop = false;
        bool availabilityBottom = false;
        bool availabilityTotal = false;
        OpenFileDialog ofd = new OpenFileDialog();
        CancellationToken _cancellationToken;

        public string wordpathfolder = " ";//лучше сохранять в папку экселя
        public string excelpathfolder = " ";
        private void wordpathbutton_Click(object sender, RoutedEventArgs e)
        {
            ofd.Filter = "Word Documents|* .doc; *.docx";
            if (ofd.ShowDialog() == true)
            {
                wordpath.Text = (ofd.FileName);
                wordpathfolder = System.IO.Path.GetDirectoryName(ofd.FileName);
                availabilityTop = true;
                checkTextbox(availabilityBottom);
            }
        }

        private void excelpathbutton_Click(object sender, RoutedEventArgs e)
        {
            ofd.Filter = "Excel Worksheets| *.xls; *.xlsx";
            if (ofd.ShowDialog() == true)
            {
                excelpath.Text = (ofd.FileName);
                excelpathfolder = System.IO.Path.GetDirectoryName(ofd.FileName);
                availabilityBottom = true;
                checkTextbox(availabilityTop);
            }
        }

        public async void executebutton_Click(object sender, RoutedEventArgs e)
        {
            byte[] wordBytes = default(byte[]);
            byte[] excBytes = default(byte[]);

            using (StreamReader sr = new StreamReader(wordpath.Text))
            {
                using (var mem = new MemoryStream())
                {
                    sr.BaseStream.CopyTo(mem);
                    wordBytes = mem.ToArray();
                }
            }

            using (StreamReader sr = new StreamReader(excelpath.Text))
            {
                using (var mem = new MemoryStream())
                {
                    sr.BaseStream.CopyTo(mem);
                    excBytes = mem.ToArray();
                }
            }

            var stream = new MemoryStream();

            await Executor.ExecuteAsync(wordBytes, excBytes, stream);
        }

        public void checkTextbox(bool checkVariable)
        {
            if (checkVariable)
            {
                executebutton.IsEnabled = true;
            }
            else
            {
                executebutton.IsEnabled = false;
            }
        }
    }
}
