using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using FNSDBAplications.connection;

namespace FNSDBAplications
{
    /// <summary>
    /// Логика взаимодействия для AdminUser.xaml
    /// </summary>
    public partial class AdminUser : Window
    {
        public AdminUser()
        {
            InitializeComponent();

            FillDataGrid();

        }
        public string login;
        public string password;
        private void FillDataGrid()
        {
                string CmdString;
                CmdString = "SELECT id, [middlename] as Фамилия,[name] as Имя,[lastname] as Отчество,[isenabled] ,[createdate],[login] as Логин,[password] as Пароль,[position] as Должность,[is_setting_allow] FROM [dbo].[User]";
                SqlCommand cmd = new SqlCommand(CmdString, Connect.cnn);
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable("dbo.User");
                sda.Fill(dt);
                datagridAdmin.ItemsSource = dt.DefaultView;
        }
        private void Insert_Click(object sender, RoutedEventArgs e)
        {
            Insert_procedure();
            FillDataGrid();
        }
        private void Insert_procedure()
        {
            Connect.cnn.Open();

            SqlCommand cmd = new SqlCommand(@"Insert into [dbo].[User] (name,middlename,lastname,isenabled,createdate,login,password,position,[is_setting_allow])values(@n,@mn,@ln,@is,@cd,@l,@p,@d,@iss)", Connect.cnn);

            cmd.Parameters.AddWithValue("@n", Name.Text);
            cmd.Parameters.AddWithValue("@ln", LastName.Text);
            cmd.Parameters.AddWithValue("@mn", MiddleName.Text);
            cmd.Parameters.AddWithValue("@is", IsEnabled.IsChecked);
            cmd.Parameters.AddWithValue("@cd", Convert.ToDateTime(CreateDate.Text));
            cmd.Parameters.AddWithValue("@l", Login.Text);
            cmd.Parameters.AddWithValue("@p", Password.Text);
            cmd.Parameters.AddWithValue("@d", Position.Text);
            cmd.Parameters.AddWithValue("@iss", Setting.IsChecked);

            cmd.ExecuteNonQuery();

            Connect.cnn.Close();

            Name.Text = null;
            LastName.Text = null;
            MiddleName.Text = null;
            IsEnabled.IsChecked = false;
            CreateDate.Text = null;
            Login.Text = null;
            Password.Text = null;
            Position.Text = null;
            Setting.IsChecked = false;

        }
        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            Delete_procedure();
            FillDataGrid();
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
        private void Delete_procedure()
        {
            if (datagridAdmin.SelectedItem == null)
                return;
            DataRowView rowView = (DataRowView)datagridAdmin.SelectedItem; // Assuming that you are having a DataTable.DefaultView as ItemsSource;
            //Подключение
            Connect.cnn.Open();
            //Процедура удаления
            SqlCommand cmd = new SqlCommand($"DELETE FROM [dbo].[User] WHERE id ={rowView["id"]}", Connect.cnn);
            cmd.ExecuteNonQuery();
            Connect.cnn.Close();

        }
        private void Update_procedure()
        {
            if (datagridAdmin.SelectedItem == null)
                 return;
            DataRowView rowView = (DataRowView)datagridAdmin.SelectedItem;
             //Подключение                                                         
            Connect.cnn.Open();
           
            string @name = "'" + rowView["Имя"] + "'";
            string @middlename = "'" + rowView["Фамилия"] + "'";
            string @lastname = "'" + rowView["Отчество"] + "'";
            string @isenabled = "'" + rowView["isenabled"] + "'";
            string @login = "'" + rowView["Логин"] + "'";
            string @password = "'" + rowView["Пароль"] + "'";
            string @position = "'" + rowView["Должность"] + "'";
            string @is_setting_allow = "'" + rowView["is_setting_allow"] + "'";

            //Процедура редактирования
            SqlCommand cmd = new SqlCommand($"UPDATE [dbo].[User] SET name = @name, middlename = @middlename,lastname = @lastname,isenabled = @isenabled," +
                $"login = @login,password = @password,position = @position,[is_setting_allow]=@is_setting_allow  WHERE id ={rowView["id"]}", Connect.cnn);
            cmd.Parameters.AddWithValue("@name", @name.Replace("'",""));
            cmd.Parameters.AddWithValue("@middlename", @middlename.Replace("'", ""));
            cmd.Parameters.AddWithValue("@lastname", @lastname.Replace("'", ""));
            cmd.Parameters.AddWithValue("@login", @login.Replace("'", ""));
            cmd.Parameters.AddWithValue("@password", @password.Replace("'", ""));
            cmd.Parameters.AddWithValue("@position", @position.Replace("'", ""));
            cmd.Parameters.AddWithValue("@isenabled", @isenabled.Replace("'", ""));
            cmd.Parameters.AddWithValue("@is_setting_allow", @is_setting_allow.Replace("'", ""));


            cmd.ExecuteNonQuery();
            Connect.cnn.Close();
        }
        private void Update_Click(object sender, RoutedEventArgs e)
        {
            Update_procedure();
            FillDataGrid();
        }
    }
}
