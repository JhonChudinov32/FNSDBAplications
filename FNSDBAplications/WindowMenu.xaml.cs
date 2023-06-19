


using FNSDBAplications.microsoft;
using Microsoft.Win32;
using System;
using System.Windows;

namespace FNSDBAplications
{
    /// <summary>
    /// Логика взаимодействия для WindowMenu.xaml
    /// </summary>
    public partial class WindowMenu : Window
    {
        public WindowMenu()
        {
            InitializeComponent();
        }
        public string login;
        public string password;

        private void UpdateDB_Click(object sender, RoutedEventArgs e)
        {
            MainWindow m = new MainWindow
            {
                login = login,
                password = password
            };
            m.Show();
            this.Hide();
        }

        private void Admin_Click(object sender, RoutedEventArgs e)
        {
            AdminUser m = new AdminUser
            {
                login = login,
                password = password
            };
            m.Show();
            this.Hide();
        }

        private void Db_Click(object sender, RoutedEventArgs e)
        {
            Fnsdb m = new Fnsdb
            {
                login = login,
                password = password
            };
            m.Show();
            this.Hide();
        }

        private void Girbo_Click(object sender, RoutedEventArgs e)
        {
            GIRBO m = new GIRBO
            {
                login = login,
                password = password
            };
            m.Show();
            this.Hide();
        }

        private void ToExcelCredit_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Excel Documents (*.xls)|*.xls",
                FileName = "Выгрузка доходы" + " " + DateTime.Today.Date.ToString("d") + ".xls"
            };
            if (sfd.ShowDialog() == true)
            {
                ImportExcel.ToExcel_Debet_Credet(sfd.FileName);
            }
        }

        private void ToExcelNalog_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Excel Documents (*.xls)|*.xls",
                FileName = "Выгрузка налоги" + " " + DateTime.Today.Date.ToString("d") + ".xls"
            };
            if (sfd.ShowDialog() == true)
            {
                ImportExcel.ToExcel_Nalog(sfd.FileName);
            }
        }

        private void ToExcelSCHR_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Excel Documents (*.xls)|*.xls",
                FileName = "Выгрузка СЧР" + " " + DateTime.Today.Date.ToString("d") + ".xls"
            };
            if (sfd.ShowDialog() == true)
            {
                ImportExcel.ToExcel_SCHR(sfd.FileName);
            }
        }

        private void ToExcelShtraf_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Excel Documents (*.xls)|*.xls",
                FileName = "Выгрузка Штрафы" + " " + DateTime.Today.Date.ToString("d") + ".xls"
            };
            if (sfd.ShowDialog() == true)
            {
                ImportExcel.ToExcel_Shtraf(sfd.FileName);
            }

        }

        private void ToExcelPeni_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Excel Documents (*.xls)|*.xls",
                FileName = "Выгрузка Пени" + " " + DateTime.Today.Date.ToString("d") + ".xls"
            };
            if (sfd.ShowDialog() == true)
            {
                ImportExcel.ToExcel_Peni(sfd.FileName);
            }
        }

        private void ToExcelRMSP_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Excel Documents (*.xls)|*.xls",
                FileName = "Выгрузка РМСП" + " " + DateTime.Today.Date.ToString("d") + ".xls"
            };
            if (sfd.ShowDialog() == true)
            {
                ImportExcel.ToExcel_RMSP(sfd.FileName);
            }
        }

        private void ToExcelUSN_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Excel Documents (*.xls)|*.xls",
                FileName = "Выгрузка УСН" + " " + DateTime.Today.Date.ToString("d") + ".xls"
            };
            if (sfd.ShowDialog() == true)
            {
                ImportExcel.ToExcel_USN(sfd.FileName);
            }
        }
    }
}
