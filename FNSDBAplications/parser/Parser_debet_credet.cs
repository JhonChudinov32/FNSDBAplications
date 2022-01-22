using System;
using System.Data.SqlClient;
using System.Windows;
using System.Xml;
using FNSDBAplications.connection;

namespace FNSDBAplications.parser
{
    public class Parser_debet_credet
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
                    String arrival = "";
                    String expense = "";

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
                                XmlNode attr3 = childnodess.Attributes.GetNamedItem("СумДоход");
                                if (attr3 != null)
                                    arrival = attr3.Value;
                                XmlNode attr4 = childnodess.Attributes.GetNamedItem("СумРасход");
                                if (attr4 != null)
                                    expense = attr4.Value;

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
                            SqlCommand cmd = new SqlCommand(@"INSERT INTO dbo.Profit(signDate,inn,arrival,expense) Values (@attr,@attr2,@attr3,@attr4)", Connect.cnn);

                            cmd.Parameters.AddWithValue("@attr", signDate);
                            cmd.Parameters.AddWithValue("@attr2", inn);
                            cmd.Parameters.AddWithValue("@attr3", arrival);
                            cmd.Parameters.AddWithValue("@attr4", expense);

                            cmd.ExecuteNonQuery();

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
