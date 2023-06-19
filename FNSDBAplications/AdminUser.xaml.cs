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
                string CmdString = string.Empty;
                CmdString = "SELECT id, [middlename] as Фамилия,[name] as Имя,[lastname] as Отчество,[isenabled] ,[createdate],[login] as Логин,[password] as Пароль,[dolshnost] as Должность,[is_setting_allow] FROM [fns].[dbo].[User]";
                SqlCommand cmd = new SqlCommand(CmdString, Connect.cnn);
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable("dbo.User");
                sda.Fill(dt);
                datagridAdmin.ItemsSource = dt.DefaultView;
        }
        //Процедура Добавления
        private void Insert_Click(object sender, RoutedEventArgs e)
        {
            Connect.cnn.Open();
            try
            {
                SqlCommand cmd = new SqlCommand(@"Insert into [fns].[dbo].[User] (name,middlename,lastname,isenabled,createdate,login,password,dolshnost,[is_setting_allow])values(@n,@mn,@ln,@is,@cd,@l,@p,@d,@iss)", Connect.cnn);
                cmd.Parameters.AddWithValue("@n", Name.Text);
                cmd.Parameters.AddWithValue("@ln", LastName.Text);
                cmd.Parameters.AddWithValue("@mn", MiddleName.Text);
                cmd.Parameters.AddWithValue("@is", IsEnabled.IsChecked);
                cmd.Parameters.AddWithValue("@cd", CreateDate.Text);
                cmd.Parameters.AddWithValue("@l", Login.Text);
                cmd.Parameters.AddWithValue("@p", Password.Text);
                cmd.Parameters.AddWithValue("@d", Position.Text);
                cmd.Parameters.AddWithValue("@iss", Setting.IsChecked);
                cmd.ExecuteNonQuery();
            }
            catch (Exception ei)
            {
                MessageBox.Show(ei.Message);
            }
            finally
            {
                Connect.cnn.Close();
                FillDataGrid();
            }
           

            

        }
        //Процедура удаление
        private void Delete_procedure()
        {
            Connect.cnn.Open();
            try
            {
                if (datagridAdmin.SelectedItem == null)
                    return;
                DataRowView rowView = (DataRowView)datagridAdmin.SelectedItem; // Assuming that you are having a DataTable.DefaultView as ItemsSource;

                //Процедура удаления
                SqlCommand cmd = new SqlCommand($"DELETE FROM [fns].[dbo].[User] WHERE id ={rowView["id"]}", Connect.cnn);
                cmd.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                Connect.cnn.Close();
            }
        }
        //Процедура обновления
        private void Update_procedure()
        {
            Connect.cnn.Open();
            try
            {
                if (datagridAdmin.SelectedItem == null)
                    return;
                DataRowView rowView = (DataRowView)datagridAdmin.SelectedItem;
                //Подключение                                                         

                //Процедура редактирования
                SqlCommand cmd = new SqlCommand($"UPDATE [fns].[dbo].[User] SET name = {rowView["name"]} , middlename = {rowView["middlename"]},lastname = {rowView["lastname"]},isenabled = {rowView["isenabled"]},createdate = {rowView["createdate"]},login = {rowView["login"]},password = {rowView["password"]},dolshnost = {rowView["dolshnost"]},[is_setting_allow]={rowView["is_setting_allow"]}  WHERE id ={rowView["id"]}", Connect.cnn);
                cmd.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                Connect.cnn.Close();
            }

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
        private void Update_Click(object sender, RoutedEventArgs e)
        {
            Update_procedure();
            FillDataGrid();
        }
     /*   private void Select_Data()
        {
            if (datagridAdmin.SelectedItem == null)
                return;
            DataRowView dr = (DataRowView)datagridAdmin.SelectedItem;
            Name.Text = (string)dr["name"];
            MiddleName.Text = (string)dr["middlename"];
            LastName.Text = (string)dr["lastname"];
            IsEnabled.IsChecked = (bool)dr["isenabled"];
            Setting.IsChecked = (bool)dr["is_setting_allow"];
            Position.Text = (string)dr["dolshnost"];
            Login.Text = (string)dr["login"];
            Password.Text = (string)dr["password"];
            CreateDate.Text = (string)dr["createdate"];

            FillDataGrid();
        }*/
    }
}
