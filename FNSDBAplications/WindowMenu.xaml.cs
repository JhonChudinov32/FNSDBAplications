
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
          //  UpdateDB.Content = new TextBlock() { Text = "Загрузка архивов открытых данных ФНС", TextWrapping = TextWrapping.Wrap };
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

        private void Company_Click(object sender, RoutedEventArgs e)
        {
            Company m = new Company
            {
                login = login,
                password = password
            };
            m.Show();
            this.Hide();
        }
    }
}
