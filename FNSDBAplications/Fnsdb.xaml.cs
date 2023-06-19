using FNSDBAplications.microsoft;
using Microsoft.Win32;
using System;
using System.Data.SqlClient;
using System.Windows;

namespace FNSDBAplications
{
    /// <summary>
    /// Логика взаимодействия для Fnsdb.xaml
    /// </summary>
    public partial class Fnsdb : Window
    {
        string CmdString = "Select dbo.SpiskiRIC.РИЦ, dbo.SpiskiRIC.Наименование,dbo.SpiskiRIC.ИНН, dbo.FnsMSP.dateMSP as [Дата включения в РСМП],  CASE FnsMSP.categoryMSP WHEN '1' THEN '1 (микропредприятие)' WHEN '2' THEN '2 (малое предприятие)' WHEN '3' THEN '3 (среднее предприятие)' ELSE 'нет информации' END AS Категория,dbo.Dochod.summDochod as Доход,dbo.Dochod.summRaschod as Расход, dbo.SSCHR.colRab as СЧР, dbo.USN.usn as УСН From dbo.SpiskiRIC left join dbo.FnsMSP on dbo.SpiskiRIC.ИНН = dbo.FnsMSP.inn left join dbo.Dochod on dbo.SpiskiRIC.ИНН = dbo.Dochod.inn left join dbo.SSCHR on dbo.SpiskiRIC.ИНН = dbo.SSCHR.inn left join dbo.USN on dbo.SpiskiRIC.ИНН = dbo.USN.inn";
        string ConString = @"Data Source=DPO-STAT1\SQLEXPRESS;Initial Catalog=fns;Integrated Security=True";
        public Fnsdb()
        {
            InitializeComponent();
            FillDataGrid();
        }
        public string login;
        public string password;
        private void FillDataGrid()
        {
            using (SqlConnection con = new SqlConnection(ConString))
            {
                SqlCommand cmd = new SqlCommand(CmdString, con);
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                System.Data.DataTable dt = new System.Data.DataTable("dbo.SpiskiRIC");
                sda.Fill(dt);
                datagrid1.ItemsSource = dt.DefaultView;
            }
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
        private void Print_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Excel Documents (*.xls)|*.xls",
                FileName = "Данные_ФНС_РИЦ" + " " + DateTime.Today.Date.ToString("d") + ".xls"
                //FileName = "Выгрузка Минцифры" + " " + DateTime.Today.Date.ToString("d") + ".docx"
            };
            if (sfd.ShowDialog() == true)
            {
                ImportExcel.ToExcel(sfd.FileName, CmdString);
            }
        }
    }
}
