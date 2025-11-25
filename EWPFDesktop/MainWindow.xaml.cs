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
        OpenFolderDialog fbd = new OpenFolderDialog();
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
            byte[] wordBytes = default;
            byte[] excBytes = default;

            Executor.ReadFileBytes(ref wordBytes, wordpath.Text);
            Executor.ReadFileBytes(ref excBytes, excelpath.Text);

            var stream = new MemoryStream();

            try
            {
                var result = await Executor.ExecuteAsync(wordBytes, excBytes, stream);

                if (ExcelRadio.IsChecked == true)
                {
                    System.IO.File.WriteAllBytes($@"{excelpathfolder}\out_file.docx", stream.ToArray());
                }
                else if (WordRadio.IsChecked == true)
                {
                    System.IO.File.WriteAllBytes($@"{wordpathfolder}\out_file.docx", stream.ToArray());
                }
                else if ((PathRadio.IsChecked == true) && (!(String.IsNullOrEmpty(OutfilePathText.Text))))
                {
                    System.IO.File.WriteAllBytes($@"{OutfilePathText.Text}\out_file.docx", stream.ToArray());
                }
                MessageBox.Show("Обработка успешно завершена");
            }
            catch
            {
                MessageBox.Show("Во время работы программы произошла ошибка. Файлы не были обработаны");
            }
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

        public void PathRadioChecked(object sender, RoutedEventArgs e)
        {
            OutfilePathText.IsEnabled = true;
        }

        public void PathRadioUnchecked(object sender, RoutedEventArgs e)
        {
            OutfilePathText.IsEnabled = false;
        }

        private void OutfilePathButton_Click(object sender, RoutedEventArgs e)
        {
            PathRadio.IsChecked = true;
            OutfilePathText.IsEnabled = true;
            if (fbd.ShowDialog() == true)
            {
                OutfilePathText.Text = fbd.FolderName;
            }
        }
    }
}
