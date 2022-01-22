using FNSDBAplications.connection;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Controls;

namespace FNSDBAplications
{
    /// <summary>
    /// Логика взаимодействия для Fnsdb.xaml
    /// </summary>
    public partial class Fnsdb : System.Windows.Window
    {
        public Fnsdb()
        {
            InitializeComponent();
            FillDataGrid();
        }
        public string login;
        public string password;
        private void FillDataGrid()
        {
            
            string CmdString;
            CmdString = "Select dbo.DateFNS.id, dbo.DateFNS.companyName as Компания,dbo.DateFNS.inn, dbo.DateFNS.[Дата включения в РСМП], Категория, " +
                "dbo.DateFNS.[Доход],dbo.DateFNS.[Расход], dbo.DateFNS.[СЧР], dbo.DateFNS.[УСН] From dbo.DateFNS";
            SqlCommand cmd = new SqlCommand(CmdString, Connect.cnn);
            SqlDataAdapter sda = new SqlDataAdapter(cmd);
            System.Data.DataTable dt = new System.Data.DataTable("dbo.DateFNS");
            sda.Fill(dt);
            dataFNS.ItemsSource = dt.DefaultView;
        }
        private void Print_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog importExcel = new SaveFileDialog
            {
                Filter = "Excel Documents (*.xls)|*.xls",
                FileName = "Выгрузка данных" + " " + DateTime.Today.Date.ToString("d") + ".xls"
            };

            importExcel.ShowDialog();
            ImportExcel_procedure(importExcel.FileName);   
        }
        public void ImportExcel_procedure(string files)
        {
            Microsoft.Office.Interop.Excel.Application xlApp;
            Workbook xlWorkBook;
            Worksheet xlWorkSheet;
            Range xlRange;

            xlApp = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = true
            };
            xlWorkBook = xlApp.Workbooks.Add();
            xlWorkBook.SaveAs(files);
            xlWorkSheet = (Worksheet)xlWorkBook.Sheets[1];
            xlWorkSheet.Columns.AutoFit();

            for (int j = 0; j < dataFNS.Columns.Count; j++)
            {
                Range myRange = (Range)xlWorkSheet.Cells[1, j + 1];
                xlWorkSheet.Cells[1, j + 1].Font.Bold = true;
                xlWorkSheet.Columns[j + 1].ColumnWidth = 15;
                myRange.Value2 = dataFNS.Columns[j].Header;
            }
            for (int i = 0; i < dataFNS.Columns.Count; i++)
            {
                for (int j = 0; j < dataFNS.Items.Count; j++)
                {
                    TextBlock b = dataFNS.Columns[i].GetCellContent(dataFNS.Items[j]) as TextBlock;
                    xlRange = (Range)xlWorkSheet.Cells[j + 2, i + 1];
                    xlRange.Value2 = b.Text;
                }
            }
            xlApp.Visible = true;
        }
        private void Back_Click(object sender, RoutedEventArgs e)
        {
            WindowMenu m=new WindowMenu
            {
                login = login,
                password = password
            };
            m.Show();
            this.Hide();
        }
    }
}
