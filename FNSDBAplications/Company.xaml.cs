using System.Data;
using System.Data.SqlClient;
using System.Windows;
using FNSDBAplications.connection;
using Microsoft.Win32;

namespace FNSDBAplications
{
    /// <summary>
    /// Логика взаимодействия для Company.xaml
    /// </summary>
    public partial class Company : Window
    {
        public Company()
        {
            InitializeComponent();
            FillDataGrid();
        }
        public string login;
        public string password;
        private void FillDataGrid()
        {
            string CmdString = string.Empty;
            CmdString = "SELECT id, [companyName] as Компания,[inn] as ИНН,[ogrn] as ОГРН,[ip] as ИП ,[location] as Адрес FROM dbo.CompanyGroup";
            SqlCommand cmd = new SqlCommand(CmdString, Connect.cnn);
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable("dbo.CompanyGroup");
            sda.Fill(dt);
            datagridComapny.ItemsSource = dt.DefaultView;
        }
        private void Closed_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
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
        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            Delete_procedure();
            FillDataGrid();
        }
        private void Update_Click(object sender, RoutedEventArgs e)
        {
            Update_procedure();
            FillDataGrid();
        }
        private void Insert_Click(object sender, RoutedEventArgs e)
        {
            Insert_procedure();
            FillDataGrid();
        }
        private void Delete_procedure()
        {
            if (datagridComapny.SelectedItem == null)
                return;
            DataRowView rowView = (DataRowView)datagridComapny.SelectedItem; // Assuming that you are having a DataTable.DefaultView as ItemsSource;
            //Подключение
            Connect.cnn.Open();
            //Процедура удаления
            SqlCommand cmd = new SqlCommand($"DELETE FROM dbo.CompanyGroup WHERE id ={rowView["id"]}", Connect.cnn);
            cmd.ExecuteNonQuery();
            Connect.cnn.Close();

        }
        private void Update_procedure()
        {
            if (datagridComapny.SelectedItem == null)
                return;
            DataRowView rowView = (DataRowView)datagridComapny.SelectedItem;
            //Подключение                                                         
            Connect.cnn.Open();

            string @companyName = "'" + rowView["Компания"] + "'";
            string @inn = "'" + rowView["ИНН"] + "'";
            string @ogrn = "'" + rowView["ОГРН"] + "'";
            string @ip = "'" + rowView["ИП"] + "'";
            string @location = "'" + rowView["Адрес"] + "'";
           

            //Процедура редактирования
            SqlCommand cmd = new SqlCommand($"UPDATE dbo.CompanyGroup SET companyName = @companyName, inn = @inn, ogrn = @ogrn, ip = @ip," +
                $"location = @location  WHERE id ={rowView["id"]}", Connect.cnn);

            cmd.Parameters.AddWithValue("@companyName", @companyName.Replace("'", ""));
            cmd.Parameters.AddWithValue("@inn", @inn.Replace("'", ""));
            cmd.Parameters.AddWithValue("@ogrn", @ogrn.Replace("'", ""));
            cmd.Parameters.AddWithValue("@ip", @ip.Replace("'", ""));
            cmd.Parameters.AddWithValue("@location", @location.Replace("'", ""));
           
            cmd.ExecuteNonQuery();
            Connect.cnn.Close();
        }
        private void Insert_procedure()
        {
            Connect.cnn.Open();

            SqlCommand cmd = new SqlCommand(@"Insert into fns.dbo.CompanyGroup (companyName, inn, ogrn, ip, location) values(@companyName, @inn, @ogrn, @ip, @location)", Connect.cnn);

            cmd.Parameters.AddWithValue("@companyName", CompanyName.Text);
            cmd.Parameters.AddWithValue("@inn", inn.Text);
            cmd.Parameters.AddWithValue("@ogrn", ogrn.Text);
            cmd.Parameters.AddWithValue("@ip", ip.Text);
            cmd.Parameters.AddWithValue("@location", location.Text);

            cmd.ExecuteNonQuery();

            Connect.cnn.Close();

            CompanyName.Text = null;
            inn.Text = null;
            ogrn.Text = null;
            ip.Text = null;
            location.Text = null;

        }
        private void ExportExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Excel файлы|*.xls;*.xlsx"
            };
            openFile.ShowDialog();
            var Filenames = openFile.FileNames;
            //цикл открытия файлов поочередно
            foreach (string fil in openFile.FileNames)
            {
                //ProgressBar
                ExportExcel_procedure(fil);
                //Parser
                FillDataGrid();
            }

        }
        private void ExportExcel_procedure(string file)
        {
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            Microsoft.Office.Interop.Excel.Range xlRange;

            int xlRow;
            Connect.cnn.Open();
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(file);
            xlWorkSheet = xlWorkBook.Worksheets["Company"];
            xlRange = xlWorkSheet.UsedRange;
            
            for (xlRow = 2; xlRow <= xlRange.Rows.Count; xlRow++)
            {
                SqlCommand cmd = new SqlCommand(@"Insert into fns.dbo.CompanyGroup (companyName, inn, ogrn, ip, location) values(@companyName, @inn, @ogrn, @ip, @location)", Connect.cnn);

                cmd.Parameters.AddWithValue("@companyName", xlRange.Cells[xlRow, 1].text);
                cmd.Parameters.AddWithValue("@inn", xlRange.Cells[xlRow, 2].text);
                cmd.Parameters.AddWithValue("@ogrn", xlRange.Cells[xlRow, 3].text);
                cmd.Parameters.AddWithValue("@ip", xlRange.Cells[xlRow, 4].text);
                cmd.Parameters.AddWithValue("@location", xlRange.Cells[xlRow, 5].text);
                cmd.ExecuteNonQuery();
                
            }

            Connect.cnn.Close();
            xlWorkBook.Close();
            xlApp.Quit();
        }

    }
}
