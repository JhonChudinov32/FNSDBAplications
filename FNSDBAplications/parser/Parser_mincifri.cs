using FNSDBAplications.connection;
using System;
using System.Data.SqlClient;
using System.IO;
using System.Windows;


namespace FNSDBAplications.parser
{
    public class Parser_mincifri
    {
        //Парсер Минцифры
        public static void Parser(string file)
        {
            //Код парсера
            try
            {
                  using (StreamReader parser = new StreamReader(file))
                  {
                       int lineNumber = 0;
                       Connect.cnn.Open();
                       while (!parser.EndOfStream)
                       {
                            var line = parser.ReadLine();
                            if (lineNumber != 0)
                            {

                                var values = line.Split(';');
                                String FullNameOrg = values[2].ToString();
                                String INN = values[4].ToString();
                                String EGRUL = values[3].ToString();
                                DateTime DataOfAccreditation = DateTime.Parse(values[5].ToString());
                                DateTime DateOfdeposit = DateTime.Parse(values[1].ToString());
                                String SolutionNumber = values[6].ToString();
                                String Status = values[9].ToString();

                                int m = 0;
                                SqlCommand SqlProv = new SqlCommand(@"SELECT COUNT(INN) As CountTabNum FROM dbo.SpisokAccrRIC WHERE [INN]= '" + INN + "' ", Connect.cnn);
                            
                                m = (Int32)(SqlProv.ExecuteScalar());

                                 if (m != 0)
                                 {
                                  
                                        // внесение данных в БД
                                        SqlCommand cmd = new SqlCommand(@"INSERT INTO dbo.Mincifri(FullNameOrg,INN,EGRUL,DataOfAccreditation,DateOfdeposit,SolutionNumber,Status) Values (@FullNameOrg,@INN,@EGRUL,@DataOfAccreditation,@DateOfdeposit,@SolutionNumber,@Status)", Connect.cnn);
                                        cmd.Parameters.AddWithValue("@FullNameOrg", FullNameOrg);
                                        cmd.Parameters.AddWithValue("@INN", INN);
                                        cmd.Parameters.AddWithValue("@EGRUL", EGRUL);
                                        cmd.Parameters.AddWithValue("@DataOfAccreditation", DataOfAccreditation);
                                        cmd.Parameters.AddWithValue("@DateOfdeposit", DateOfdeposit);
                                        cmd.Parameters.AddWithValue("@SolutionNumber", SolutionNumber);
                                        cmd.Parameters.AddWithValue("@Status", Status);
                                        cmd.ExecuteNonQuery();

                                        //добавляем новую запись в таблицу
                                    
                                 }
                            }
                            lineNumber++;
                       }
                  }
                  
            }

            catch (Exception ex)
            {
                  MessageBox.Show(ex.Message);
            }
            finally
            {
                Connect.cnn.Close();
                MessageBox.Show("Данные выгружены");
            }
        }
  
    }
}
