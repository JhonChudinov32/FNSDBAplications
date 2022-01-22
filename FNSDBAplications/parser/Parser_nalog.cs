using System;
using System.Data.SqlClient;
using System.Windows;
using System.Xml;
using FNSDBAplications.connection;

namespace FNSDBAplications.parser
{
    public class Parser_nalog
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
                    String name = "";
                    String payment = "";

                    if (xnode.Attributes.Count > 0)
                    {
                        XmlNode attr = xnode.Attributes.GetNamedItem("ДатаСост");
                        if (attr != null)
                            signDate = DateTime.ParseExact(attr.Value, "mm.dd.yyyy", null);

                        // получаем атрибуты с первого корневого узла
                        foreach (XmlNode childnodes in xnode.ChildNodes)
                        {
                            if (childnodes.Name == "СведНП")
                            {
                                XmlNode attr2 = childnodes.Attributes.GetNamedItem("ИННЮЛ");
                                if (attr2 != null)
                                    inn = attr2.Value.ToUpper();


                                foreach (XmlNode child in xnode.SelectNodes("СвУплСумНал"))
                                {

                                    if (child.Attributes.Count > 0)
                                    {

                                        XmlNode attr3 = child.Attributes.GetNamedItem("НаимНалог");
                                        if (attr3 != null)
                                            name = attr3.Value;

                                        XmlNode attr4 = child.Attributes.GetNamedItem("СумУплНал");
                                        if (attr4 != null)
                                            payment = attr4.Value.ToUpper();

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
                                        SqlCommand cmd = new SqlCommand(@"INSERT INTO dbo.Nalog(signDate, inn,name,payment) Values (@attr,@attr2,@attr3,@attr4)", Connect.cnn);

                                        cmd.Parameters.AddWithValue("@attr", signDate);
                                        cmd.Parameters.AddWithValue("@attr2", inn);
                                        cmd.Parameters.AddWithValue("@attr3", name);
                                        cmd.Parameters.AddWithValue("@attr4", payment);

                                        cmd.ExecuteNonQuery();

                                        //добавляем новую запись в таблицу

                                        Connect.cnn.Close();
                                    }
                                }
                            }
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
