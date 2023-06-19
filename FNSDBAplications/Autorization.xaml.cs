using System;
using System.Data.SqlClient;
using System.IO;
using System.Windows;
using System.Xml.Serialization;
using FNSDBAplications.connection;
using SaveOptions = FNSDBAplications.Toolsuser.SaveOptions;

namespace FNSDBAplications
{
    /// <summary>
    /// Логика взаимодействия для Autorization.xaml
    /// </summary>
    public partial class Autorization : Window
    {
        
        private SaveOptions saveOptions;
        public Autorization()
        {
            if (saveOptions == null)
            {
                // Определяем, существует ли указанный файл
                if (File.Exists("SaveOptions.xml"))
                {
                    using (var stream = File.OpenRead("SaveOptions.xml"))
                    {
                        var serializer = new XmlSerializer(typeof(SaveOptions));
                        saveOptions = serializer.Deserialize(stream) as SaveOptions;
                    }
                }
                else
                    saveOptions = new SaveOptions();
            }
            DataContext = saveOptions;
            InitializeComponent();
        }
        private void Entered_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Connect.cnn.Open();
                int k = 0;

                SqlCommand cmd = new SqlCommand("SELECT COUNT(middlename)as qw FROM [fns].[dbo].[User] WHERE [login] ='" + Login.Text + "' and [password] ='" + Password.Password + "' and [isenabled] = 'true' ", Connect.cnn);

                k = (Int32)cmd.ExecuteScalar();

                if (Login.Text == "" || Password.Password == "") // Определить, является ли вход пустым
                {
                    Connect.cnn.Close();
                    MessageBox.Show("Пожалуйста, введите имя пользователя и пароль");

                    using (var stream = File.Open("SaveOptions.xml", FileMode.Create))
                    {
                        if (LoginIDMemory.IsChecked == false)
                        {
                            saveOptions.SaveLoginID = "";
                            saveOptions.SaveLoginPSW = "";
                        }
                        var serializer = new XmlSerializer(typeof(SaveOptions));
                        serializer.Serialize(stream, saveOptions);
                    }
                }

                else
                {
                    if (k != 0) // определяет, есть ли имя пользователя и пароль, введенные пользователем
                    {
                        //MessageBox.Show("Вход в систему успешен"); // Показать главное окно;
                        
                        using (var stream = File.Open("SaveOptions.xml", FileMode.Create))
                        {
                            if (LoginIDMemory.IsChecked == false)
                            {
                                saveOptions.SaveLoginID = "";
                                saveOptions.SaveLoginPSW = "";
                            }
                            var serializer = new XmlSerializer(typeof(SaveOptions));
                            serializer.Serialize(stream, saveOptions);
                            WindowMenu m = new WindowMenu
                            {
                                login = this.Login.Text,
                                password = this.Password.Password
                            };
                            m.Show();
                            this.Hide();

                        }
                       
                        Connect.cnn.Close(); // Закрыть соединение с базой данных
                        this.Close();
                    }

                    else
                    {
                        MessageBox.Show("Неверное имя пользователя или пароль");
                        using (var stream = File.Open("SaveOptions.xml", FileMode.Create))
                        {
                            if (LoginIDMemory.IsChecked == false)
                            {
                                saveOptions.SaveLoginID = "";
                                saveOptions.SaveLoginPSW = "";
                            }
                            var serializer = new XmlSerializer(typeof(SaveOptions));
                            serializer.Serialize(stream, saveOptions);
                        }
                        Connect.cnn.Close(); // Закрыть соединение с базой данных
                    }
                }
            }
            catch (Exception er)
            {
                MessageBox.Show(er.ToString());
            }
           
        }
        private void Closed_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }
        private void SaveOptionsXml()
        {
            using (var stream = File.Open("SaveOptions.xml", FileMode.Create))
            {
                if (LoginIDMemory.IsChecked == false)
                {
                    saveOptions.SaveLoginID = "";
                    saveOptions.SaveLoginPSW = "";
                }
                var serializer = new XmlSerializer(typeof(SaveOptions));
                serializer.Serialize(stream, saveOptions);
            }
        }
    }
}
