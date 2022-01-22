using FNSDBAplications.connection;
using Microsoft.Win32;
using System.Windows;
using FNSDBAplications.parser;
using System.Threading;
using System.Windows.Threading;
using System;

namespace FNSDBAplications
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
        public string login;
        public string password;
        private void Button_Click(object sender, RoutedEventArgs e)
        {
             OpenFileDialog openFile = new OpenFileDialog
                {
                    //SaveFileDialog saveFile = new SaveFileDialog();
                    Multiselect = true,
                Filter = "XML файлы|*.xml"
                };
                // saveFile.Filter = "XML файлы|*.xml";

                //string[] buf;
                openFile.ShowDialog();
            
          //  buf = File.ReadAllLines(openFile.FileName, Encoding.Default);
          //  saveFile.ShowDialog();
          //  File.WriteAllLines(saveFile.FileName, buf, Encoding.Default);

            var Filenames = openFile.FileNames;
            //цикл открытия файлов поочередно
            foreach (string fil in openFile.FileNames)
            {
                //ProgressBar
                Loadfiles(Filenames);
                //Parser
                Parser_rmsp.Parser(fil);
            }
        }
        public void Loadfiles(string[] filenames)
        {
              ProgBar.Maximum = filenames.Length;
              ProgBar.Value = ProgBar.Value + 1;
              double gg = ProgBar.Maximum / 100D;
              double percent = ProgBar.Value / gg;
            Count.Content = "Количество:" + ProgBar.Value.ToString() + "/" + ProgBar.Maximum.ToString();
            Percent.Content = string.Format("{0:0.##}", percent) +"%";
            Application.Current.Dispatcher.Invoke(DispatcherPriority.Background,
                                         new Action(delegate { }));
        }

        private void ProgressBar_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {

        }

        private void DebetCredet_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "XML файлы|*.xml"
            };
            openFile.ShowDialog();
            var Filenames = openFile.FileNames;
            //цикл открытия файлов поочередно
            foreach (string fil in openFile.FileNames)
            {
                //ProgressBar
                Loadfiles(Filenames);
                //Parser
                Parser_debet_credet.Parser(fil);
            }
        }

        private void Ckr_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "XML файлы|*.xml"
            };
            openFile.ShowDialog();
            var Filenames = openFile.FileNames;
            //цикл открытия файлов поочередно
            foreach (string fil in openFile.FileNames)
            {
                //ProgressBar
                Loadfiles(Filenames);
                //Parser
                Parser_scr.Parser(fil);
            }
        }

        private void Nalog_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "XML файлы|*.xml"
            };
            openFile.ShowDialog();
            var Filenames = openFile.FileNames;
            //цикл открытия файлов поочередно
            foreach (string fil in openFile.FileNames)
            {
                //ProgressBar
                Loadfiles(Filenames);
                //Parser
                Parser_nalog.Parser(fil);
            }
        }

        private void Shtraf_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "XML файлы|*.xml"
            };
            openFile.ShowDialog();
            var Filenames = openFile.FileNames;
            //цикл открытия файлов поочередно
            foreach (string fil in openFile.FileNames)
            {
                //ProgressBar
                Loadfiles(Filenames);
                //Parser
                Parser_shtraf.Parser(fil);
            }
        }

        private void Peni_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "XML файлы|*.xml"
            };
            openFile.ShowDialog();
            var Filenames = openFile.FileNames;
            //цикл открытия файлов поочередно
            foreach (string fil in openFile.FileNames)
            {
                //ProgressBar
                Loadfiles(Filenames);
                //Parser
                Parser_peni.Parser(fil);
            }
        }

        private void Usn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "XML файлы|*.xml"
            };
            openFile.ShowDialog();
            var Filenames = openFile.FileNames;
            //цикл открытия файлов поочередно
            foreach (string fil in openFile.FileNames)
            {
                //ProgressBar
                Loadfiles(Filenames);
                //Parser
                Parser_usn.Parser(fil);
            }
        }

        private void Back_Click(object sender, RoutedEventArgs e)
        {
            WindowMenu m = new WindowMenu
            {
                login = login,
                password = password
            };
            m.Show();
            this.Hide();
        }

        private void Closed_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }
    }
}
