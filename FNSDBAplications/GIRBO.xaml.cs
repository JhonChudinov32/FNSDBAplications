using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using FNSDBAplications.parser;
using FNSDBAplications.microsoft;
using Microsoft.Win32;
using FNSDBAplications.connection;

namespace FNSDBAplications
{
    /// <summary>
    /// Логика взаимодействия для GIRBO.xaml
    /// </summary>
    public partial class GIRBO : Window
    {
        public GIRBO()
        {
            InitializeComponent();
            FillDataGrid();
        }
        public string login;
        public string password;

        private void FillDataGrid()
        {
            string CmdString = "SELECT [РИЦ],[Наименование],[ИНН],[БАЛАНС],[Финансовый результат],[Год] FROM [fns].[dbo].[GIRBO]";
            Connect.cnn.Open();
            SqlCommand cmd = new SqlCommand(CmdString, Connect.cnn);
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable("dbo.GIRBO");
            sda.Fill(dt);
            girbogrid.ItemsSource = dt.DefaultView;
            Connect.cnn.Close();

        }
        private void Balance_Click(object sender, RoutedEventArgs e)
        {
            string inn;
            int year;
            try
            {
                if (girbogrid.SelectedItem == null)
                    return;
                DataRowView rowView = (DataRowView)girbogrid.SelectedItem; 

                inn = rowView["ИНН"].ToString();
                year = (int)rowView["Год"];
                Parser_Girbo.Load_PDF_GIRBO_Balance(inn, year);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
          
        }
        private void Finotchet_Click(object sender, RoutedEventArgs e)
        {
            string inn;
            int year;
            try
            {
                if (girbogrid.SelectedItem == null)
                    return;
                DataRowView rowView = (DataRowView)girbogrid.SelectedItem;

     
                inn = rowView["ИНН"].ToString();
                year = (int)rowView["Год"];
                Parser_Girbo.Load_PDF_GIRBO_Financial(inn, year);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        private void LoadPDF_Click(object sender, RoutedEventArgs e)
        {
            string Folder = @"X:\ДПО\Программирование\FNSDBAplications\temp\ГИРБО";
            string[] parsfiles = System.IO.Directory.GetFiles(Folder);
            foreach (string files in parsfiles)
            {
               // Loadfiles(parsfiles);
                Parser_Girbo.ParselineSQL(files);
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
        private void Excel_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Excel Documents (*.xls)|*.xls",
                FileName = "Выгрузка ГИРБО" + " " + DateTime.Today.Date.ToString("d") + ".xls"
            };
            if (sfd.ShowDialog() == true)
            {
                ImportExcel.ToExcel_Girbo(sfd.FileName);
            }
        }
    }
}
