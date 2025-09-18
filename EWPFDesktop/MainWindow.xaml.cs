using Microsoft.Win32;
using RabbitMQ.Client;
using System;
using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Shapes;

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
        //Forms.FolderBrowserDialog fbd = new Forms.FolderBrowserDialog();

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

        public async Task executebutton_Click(object sender, RoutedEventArgs e)
        {
            ConnectionFactory factory = new();
            factory.Uri = new Uri(uriString: "amqp://guest:guest@localhost:1011");// add appsettings
            factory.ClientProvidedName = "EW filebytes sender app";

            var cnn = await factory.CreateConnectionAsync();
            var channel = await cnn.CreateChannelAsync();

            string exchangeName = "EWPFExchange";
            string routingKey = "ew-routing-key";
            string queueName = "EWFileQueue";

            channel.ExchangeDeclareAsync(exchangeName, ExchangeType.Direct);
            channel.QueueDeclareAsync(queueName, durable: false, exclusive: false, autoDelete: false, arguments: null);
            channel.QueueBindAsync(queueName, exchangeName, routingKey, arguments: null);

            byte[] wordBytes = default(byte[]);
            byte[] excDoc = default(byte[]);

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
                    excDoc = mem.ToArray();
                }
            }

            //var fileBatch = channel.CreateBasicPublishBatch();
            await channel.BasicPublishAsync(exchangeName, routingKey, true, wordBytes, _cancellationToken);// need to roll into one?
            await channel.BasicPublishAsync(exchangeName, routingKey, true, excDoc, _cancellationToken);
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
