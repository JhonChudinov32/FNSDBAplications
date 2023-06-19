using System;
using System.Windows;
using System.Windows.Threading;
using FNSDBAplications.parser;

namespace FNSDBAplications
{
    /// <summary>
    /// Логика взаимодействия для Progress.xaml
    /// </summary>
    public partial class Progress : Window
    {
        public Progress(ref string[] files)
        {
            InitializeComponent();
          //  dbcd(files);
        }
       
        public void Loadfiles(string[] filenames)
        {
            ProgBar.Maximum = filenames.Length;
            ProgBar.Value = ProgBar.Value + 1;
            double gg = ProgBar.Maximum / 100D;
            double percent = ProgBar.Value / gg;
            Count.Content = "Количество:" + ProgBar.Value.ToString() + "/" + ProgBar.Maximum.ToString();
            Percent.Content = string.Format("{0:0.##}", percent) + "%";
            Application.Current.Dispatcher.Invoke(DispatcherPriority.Background,
                                         new Action(delegate { }));
        }
        public void dbcd(string[] files)
        {
            foreach (string fil in files)
            {
                Loadfiles(files);
                Parser_debet_credet.Parser(fil);
            }
        }
    }
}
