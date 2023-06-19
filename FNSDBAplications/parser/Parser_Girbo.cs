using System;
using Appl = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Windows;
using FNSDBAplications.connection;
using System.IO;
using System.Data.SqlClient;
using System.Runtime.Serialization.Formatters.Binary;
using System.Diagnostics;

namespace FNSDBAplications.parser
{
    public static class Parser_Girbo
    {
        //Конвертирование в PDF
        public static void ConvertExcelAsMemoryStream(string files)
        {
            Appl.Application app;
            Appl.Workbook workbook;
            app = new Appl.Application
            {
                Visible = false
            };

            Appl.XlFixedFormatQuality paramExportQuality =
            Appl.XlFixedFormatQuality.xlQualityStandard;
            bool paramOpenAfterPublish = false;
            bool paramIncludeDocProps = true;
            bool paramIgnorePrintAreas = true;
            object paramFromPage = Type.Missing;
            object paramToPage = Type.Missing;

            workbook = app.Workbooks.Open(files, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value); //к вашей книге
            try
            {
                for (int i = 2; i <= 3; i++)
                {
                    app.Worksheets[i].ExportAsFixedFormat(Appl.XlFixedFormatType.xlTypePDF, files + i,
                        paramExportQuality, paramIncludeDocProps, paramIgnorePrintAreas, paramFromPage, paramToPage, paramOpenAfterPublish);//куда сохраняете
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                workbook.Close();
                app.Quit();
            }
        }
        //Занесение пути ссылок в БД
        public static void Parseline(string[] files)
        {
            try
            {
                Connect.cnn.Open();

                foreach (var fil in files)
                {
                    // Loadfiles(files);
                    if (fil != null)
                    {
                        if (fil.Contains("xlsx2.pdf"))
                        {
                            string f2 = Path.GetFileName(fil);
                            int k = 0;
                            SqlCommand SqlProv;
                            SqlProv = new SqlCommand("select COUNT(SpiskiRIC.id) as CountTabNum from [Balance] right join [SpiskiRIC] on dbo.SpiskiRIC.ИНН = dbo.[Balance].inn" +
                                " where [Balance] ='" + f2 + "' ", Connect.cnn);
                            k = (Int32)SqlProv.ExecuteScalar();

                            if (k == 0)
                            {
                                SqlCommand cmd1 = new SqlCommand(@"INSERT INTO [Balance]([inn],[Balance]) Values ('" + f2[8] + f2[9] + f2[10] + f2[11] + f2[12] + f2[13] + f2[14] + f2[15] + f2[16] + f2[17] + "', @f2)", Connect.cnn);
                                cmd1.Parameters.AddWithValue("@f2", f2);
                                cmd1.ExecuteNonQuery();
                            }
                        }
                        if (fil.Contains("xlsx3.pdf"))
                        {
                            string f3 = Path.GetFileName(fil);
                            int k = 0;
                            SqlCommand SqlProv;
                            SqlProv = new SqlCommand("select COUNT(SpiskiRIC.id) as CountTabNum from [Financial Result] right join [SpiskiRIC] on dbo.SpiskiRIC.ИНН = dbo.[Financial Result].inn " +
                                "where [Financial Result] ='" + f3 + "' ", Connect.cnn);
                            k = (Int32)SqlProv.ExecuteScalar();

                            if (k == 0)
                            {
                                SqlCommand cmd = new SqlCommand(@"INSERT INTO [Financial Result]([inn],[Financial Result]) Values ('" + f3[8] + f3[9] + f3[10] + f3[11] + f3[12] + f3[13] + f3[14] + f3[15] + f3[16] + f3[17] + "',  @f3)", Connect.cnn);
                                cmd.Parameters.AddWithValue("@f3", f3);
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                Connect.cnn.Close();
                MessageBox.Show("Данные ГИРБО внесены в БД");
            }
        }
        //Распаковка архивов
        public static void ArcGirbo(string file, string inFolder)
        {
            try
            {
                string folder = inFolder;// папка, в которую надо распаковать архив
                System.Diagnostics.Process p = new System.Diagnostics.Process();
                p.StartInfo.FileName = @"C:\Program Files\WinRAR\WinRAR.exe";
                p.StartInfo.Arguments = string.Format("x -o+ \"{0}\" \"{1}\"", file, folder);
                p.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                p.EnableRaisingEvents = true;
                p.Start();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }
        //Занесение PDF в БД тип Varbinary(max)
        public static void ParselineSQL(string fil)
        {
            try
            {
                Connect.cnn.Open();

                if (fil != null)
                {
                    if (fil.Contains("xlsx2.pdf"))
                    {
                        //Считываем файл и переводим в код
                        FileStream fs = File.OpenRead(fil);
                        byte[] contenst = new byte[fs.Length];
                        fs.Read(contenst, 0, (int)fs.Length);
                        fs.Close();

                        string f2 = Path.GetFileName(fil);
                        string inn = "" + f2[8] + f2[9] + f2[10] + f2[11] + f2[12] + f2[13] + f2[14] + f2[15] + f2[16] + f2[17] + "";
                        int k = 0;
                        SqlCommand SqlProv;
                        SqlProv = new SqlCommand("select COUNT(SpiskiRIC.id) as CountTabNum from [BalanceSQL] right join [SpiskiRIC] on dbo.SpiskiRIC.ИНН = dbo.[BalanceSQL].inn" +
                            " where [Balance] ='" + contenst + "' ", Connect.cnn);
                        k = (Int32)SqlProv.ExecuteScalar();

                        if (k == 0)
                        {
                            SqlCommand cmd1 = new SqlCommand(@"INSERT INTO [BalanceSQL]([inn],[Balance]) Values (@inn, @pdf)", Connect.cnn);
                            cmd1.Parameters.AddWithValue("@inn", inn);
                            cmd1.Parameters.AddWithValue("@pdf", contenst);
                            cmd1.ExecuteNonQuery();
                        }
                    }
                    if (fil.Contains("xlsx3.pdf"))
                    {
                        //Считываем файл и переводим в код
                        FileStream fs = File.OpenRead(fil);
                        byte[] contenst = new byte[fs.Length];
                        fs.Read(contenst, 0, (int)fs.Length);
                        fs.Close();

                        string f3 = Path.GetFileName(fil);
                        string inn = "" + f3[8] + f3[9] + f3[10] + f3[11] + f3[12] + f3[13] + f3[14] + f3[15] + f3[16] + f3[17] + "";
                        int k = 0;
                        SqlCommand SqlProv;
                        SqlProv = new SqlCommand("select COUNT(SpiskiRIC.id) as CountTabNum from [FinancialSQL] right join [SpiskiRIC] on dbo.SpiskiRIC.ИНН = dbo.[FinancialSQL].inn " +
                            "where [Financial Result] ='" + contenst + "' ", Connect.cnn);
                        k = (Int32)SqlProv.ExecuteScalar();

                        if (k == 0)
                        {
                            SqlCommand cmd2 = new SqlCommand(@"INSERT INTO [FinancialSQL]([inn],[Financial Result]) Values (@inn,  @pdf)", Connect.cnn);
                            cmd2.Parameters.AddWithValue("@inn", inn);
                            cmd2.Parameters.AddWithValue("@pdf", contenst);
                            cmd2.ExecuteNonQuery();
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                Connect.cnn.Close();
                MessageBox.Show("Данные ГИРБО внесены в БД");
            }
        }
        //Считывание файла PDF из БД
        public static void Load_PDF_GIRBO_Balance(string inn, int year)
        {
            //Путь сохранения файла PDF в корень программы
            var exePath = AppDomain.CurrentDomain.BaseDirectory;
            //Переменная значения ИНН
            string INN = inn;
            int Year = year;
            //SQL запрос
            string SQL = "SELECT [Balance], [inn] from [BalanceSQL] WHERE inn=@inn and YEAR(year) = @year";
            //Переменная для записи значения типа VARBINARY
            byte[] bytes;
            //открытие подключения
            Connect.cnn.Open();
            try
            {
                SqlCommand com = new SqlCommand(SQL, Connect.cnn);
                com.Parameters.Add(new SqlParameter("@inn", INN));
                com.Parameters.Add(new SqlParameter("@year", Year));

                //Декадируем в файл PDF
                using (SqlDataReader sdr = com.ExecuteReader())
                {
                    sdr.Read();
                    bytes = (byte[])sdr["Balance"];
                    BinaryFormatter bf = new BinaryFormatter();
                    MemoryStream ms = new MemoryStream();
                    bf.Serialize(ms, bytes);
                    bytes = ms.ToArray();
                    File.WriteAllBytes(INN + ".pdf", bytes);
                    Process.Start(exePath + INN + ".pdf");
                }
            }
            catch 
            {
                MessageBox.Show("Нет данных");
            }
            finally
            {
                Connect.cnn.Close();
                
            }
        }
        public static void Load_PDF_GIRBO_Financial(string inn, int year)
        {
            //Путь сохранения файла PDF в корень программы
            var exePath = AppDomain.CurrentDomain.BaseDirectory;
            //Переменная значения ИНН
            string INN = inn;
            int Year = year;
            //SQL запрос
            string SQL = "SELECT [Financial Result], [inn] from [FinancialSQL] WHERE inn=@inn and YEAR(year) = @year";
            //Переменная для записи значения типа VARBINARY
            byte[] bytes;
            //открытие подключения
            Connect.cnn.Open();
            try
            {
                SqlCommand com = new SqlCommand(SQL, Connect.cnn);
                com.Parameters.Add(new SqlParameter("@inn", INN));
                com.Parameters.Add(new SqlParameter("@year", Year));

                //Декадируем в файл PDF
                using (SqlDataReader sdr = com.ExecuteReader())
                {
                    sdr.Read();
                    bytes = (byte[])sdr["Financial Result"];
                    BinaryFormatter bf = new BinaryFormatter();
                    MemoryStream ms = new MemoryStream();
                    bf.Serialize(ms, bytes);
                    bytes = ms.ToArray();
                    File.WriteAllBytes(INN + ".pdf", bytes);
                    Process.Start(exePath + INN + ".pdf");
                }
            }
            catch 
            {
                MessageBox.Show("Нет данных");
            }
            finally
            {
                Connect.cnn.Close();
            }
        }
    }
}
