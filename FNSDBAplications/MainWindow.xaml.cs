using Microsoft.Win32;
using System.Windows;
using FNSDBAplications.parser;
using System.Windows.Threading;
using System;
using FNSDBAplications.microsoft;



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
        private void RMSP_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "XML файлы|*.xml"
            };
            
            //string[] buf;
            openFile.ShowDialog();

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
            Progress p = new Progress(ref Filenames);
            
               
            
            p.Show();
           

            /* foreach (string fil in openFile.FileNames)
             {

                 //Loadfiles(files);
                 Parser_debet_credet.Parser(fil);
             }*/

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
        private void Mincifri_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "csv файлы|*.csv"
            };
            openFile.ShowDialog();
            var Filenames = openFile.FileNames;
            //цикл открытия файлов поочередно
            foreach (string fil in openFile.FileNames)
            {
                Loadfiles(Filenames);
                //Parser
                Parser_mincifri.Parser(fil);
            }
        }
        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Excel Documents (*.xls)|*.xls",
               FileName = "Выгрузка Минцифры" + " " + DateTime.Today.Date.ToString("d") + ".xls"
                 //FileName = "Выгрузка Минцифры" + " " + DateTime.Today.Date.ToString("d") + ".docx"
            };
            if (sfd.ShowDialog() == true)
            {
                String sql = @"SELECT [RIC] as [РИЦ],[Name] as [Наименоввание],[Status],[ITLgota],[State],[Accreditation],[INN],[StatusMinc] FROM [fns].[dbo].[RicAccreditation] ORDER by [RIC] ASC,[Status] DESC";
                ImportExcel.ToExcel(sfd.FileName, sql);
                //ImportWord.ExportWord(sfd.FileName, sql);
            }
        }
        private void Girbo_Click(object sender, RoutedEventArgs e)
        {
            string Folder = @"X:\ДПО\Программирование\FNSDBAplications\temp\ГИРБО";

            OpenFileDialog openFile = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "ZIP файлы|*.zip"
            };
            openFile.ShowDialog();

            var Filenames = openFile.FileNames;
            //цикл открытия файлов поочередно
            foreach (string fil in openFile.FileNames)
            {
                Loadfiles(Filenames);
                Parser_Girbo.ArcGirbo(fil, Folder); 
            }
            string[] convfiles = System.IO.Directory.GetFiles(Folder);
            foreach (string confil in convfiles)
            {
                Loadfiles(convfiles);
                Parser_Girbo.ConvertExcelAsMemoryStream(confil);
            }
            string[] parsfiles = System.IO.Directory.GetFiles(Folder); 
            Parser_Girbo.Parseline(parsfiles);
        }

    }
}