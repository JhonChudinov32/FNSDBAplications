using System;
using System.Data.SqlClient;
using System.Windows;
using System.Xml;
using FNSDBAplications.connection;

namespace FNSDBAplications.parser
{
    public class Parser_peni
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
                    String sumArrearsN = "";
                    String sumPeni = "";
                    String sumFine = "";
                    String countSum = "";

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
                                    XmlNode attr3 = childnodess.Attributes.GetNamedItem("НаимНалог");
                                    if (attr3 != null)
                                        name = attr3.Value.ToUpper();

                                    XmlNode attr4 = childnodess.Attributes.GetNamedItem("СумНедНалог");
                                    if (attr4 != null)
                                        sumArrearsN = attr4.Value.ToUpper();

                                    XmlNode attr5 = childnodess.Attributes.GetNamedItem("СумПени");
                                    if (attr5 != null)
                                        sumPeni = attr5.Value.ToUpper();

                                    XmlNode attr6 = childnodess.Attributes.GetNamedItem("СумШтраф");
                                    if (attr6 != null)
                                        sumFine = attr6.Value.ToUpper();

                                    XmlNode attr7 = childnodess.Attributes.GetNamedItem("ОбщСумНедоим");
                                    if (attr7 != null)
                                        countSum = attr7.Value.ToUpper();
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
                                    SqlCommand cmd = new SqlCommand(@"INSERT INTO dbo.Arrears(signDate,inn,name,sumArrearsN,sumPeni,sumFine,countSum) Values (@attr,@attr2,@attr3,@attr4,@attr5,@attr6,@attr7)", Connect.cnn);

                                    cmd.Parameters.AddWithValue("@attr", signDate);
                                    cmd.Parameters.AddWithValue("@attr2", inn);
                                    cmd.Parameters.AddWithValue("@attr3", name);
                                    cmd.Parameters.AddWithValue("@attr4", sumArrearsN);
                                    cmd.Parameters.AddWithValue("@attr5", sumPeni);
                                    cmd.Parameters.AddWithValue("@attr6", sumFine);
                                    cmd.Parameters.AddWithValue("@attr7", countSum);

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
