using System;
using System.Data.SqlClient;
using System.Windows;
using System.Xml;
using FNSDBAplications.connection;

namespace FNSDBAplications.parser
{
    public class Parser_scr
    {
        public static void Parser(string file)
        {
            //Код парсера
            try
            {

                XmlDocument xDoc = new XmlDocument();
                xDoc.Load(file);
                // получим корневой элемент
                XmlElement xRoot = xDoc.DocumentElement;

                foreach (XmlNode xnode in xRoot)
                {
                    DateTime signDate = DateTime.Now;
                    String inn = "";
                    String quantityOfE = "";

                    // получаем атрибуты с первого корневого узла
                    if (xnode.Attributes.Count > 0)
                    {
                        XmlNode attr = xnode.Attributes.GetNamedItem("ДатаСост");
                        if (attr != null)
                            signDate = DateTime.ParseExact(attr.Value, "mm.dd.yyyy", null);

                        foreach (XmlNode childnodes in xnode.ChildNodes)
                        {
                            if (childnodes.Name == "СведНП")
                            {
                                XmlNode attr2 = childnodes.Attributes.GetNamedItem("ИННЮЛ");
                                if (attr2 != null)
                                    inn = attr2.Value;
                            }

                            foreach (XmlNode childnodess in xnode.ChildNodes)

                            {
                                XmlNode attr3 = childnodess.Attributes.GetNamedItem("КолРаб");

                                if (attr3 != null)
                                {
                                    quantityOfE = attr3.Value;
                                }

                            }
                        }
                        int m = 0;
                        SqlCommand SqlProv = new SqlCommand(@"SELECT COUNT(inn) As CountTabNum FROM dbo.CompanyGroup WHERE inn= '" + inn + "' ", Connect.cnn);
                        Connect.cnn.Open();
                        m = (Int32)(SqlProv.ExecuteScalar());
                        Connect.cnn.Close();

                        if (m != 0)
                        {
                            // подключение к БД
                            Connect.cnn.Open();
                            // внесение данных в БД
                            SqlCommand cmd = new SqlCommand(@"INSERT INTO dbo.AverNumEmp(signDate,inn,quantityOfE) Values (@attr,@attr2,@attr3)", Connect.cnn);

                            cmd.Parameters.AddWithValue("@attr", signDate);
                            cmd.Parameters.AddWithValue("@attr2", inn);
                            cmd.Parameters.AddWithValue("@attr3", quantityOfE);

                            cmd.ExecuteNonQuery();

                            //добавляем новую запись в таблицу

                            Connect.cnn.Close();
                        }
                    }
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
