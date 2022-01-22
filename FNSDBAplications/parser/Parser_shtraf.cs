using System;
using System.Data.SqlClient;
using System.Windows;
using System.Xml;
using FNSDBAplications.connection;

namespace FNSDBAplications.parser
{
    public class Parser_shtraf
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
                    String sum = "";


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
                                    inn = attr2.Value.ToUpper();

                                foreach (XmlNode childnodess in xnode.ChildNodes)

                                {
                                    XmlNode attr3 = childnodess.Attributes.GetNamedItem("СумШтраф");
                                    if (attr3 != null)
                                        sum = attr3.Value.ToUpper();

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
                                    SqlCommand cmd = new SqlCommand(@"INSERT INTO dbo.Offenses(signDate,inn,sum) Values (@attr,@attr2,@attr3)", Connect.cnn);

                                    cmd.Parameters.AddWithValue("@attr", signDate);
                                    cmd.Parameters.AddWithValue("@attr2", inn);
                                    cmd.Parameters.AddWithValue("@attr3", sum);

                                    cmd.ExecuteNonQuery();

                                    //добавляем новую запись в таблицу

                                    Connect.cnn.Close();
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
