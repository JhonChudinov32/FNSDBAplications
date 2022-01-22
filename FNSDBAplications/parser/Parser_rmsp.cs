using System;
using System.Data.SqlClient;
using System.Windows;
using System.Xml;
using FNSDBAplications.connection;

namespace FNSDBAplications.parser
{
    public  class Parser_rmsp
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
                    DateTime datemsp = DateTime.Now;
                    String category = "";
                    String inn = "";
                    String name = "";
                    String nameS = "";


                    // получаем атрибуты с первого корневого узла
                    if (xnode.Attributes.Count > 0)
                    {
                        XmlNode attr = xnode.Attributes.GetNamedItem("ДатаВклМСП");
                        if (attr != null)
                            datemsp = DateTime.ParseExact(attr.Value, "mm.dd.yyyy", null);

                        XmlNode attr1 = xnode.Attributes.GetNamedItem("КатСубМСП");
                        if (attr1 != null)
                            category = attr1.Value;

                        foreach (XmlNode childnodes in xnode.ChildNodes)
                        {
                            if (childnodes.Name == "ИПВклМСП")
                            {
                                XmlNode attr5 = childnodes.Attributes.GetNamedItem("ИННФЛ");
                                if (attr5 != null)
                                    inn = attr5.Value;
                            }
                            foreach (XmlNode childnodess in childnodes.SelectNodes("ФИОИП"))

                            {
                                XmlNode at1 = childnodess.Attributes.GetNamedItem("Фамилия");
                                XmlNode at2 = childnodess.Attributes.GetNamedItem("Имя");
                                XmlNode at3 = childnodess.Attributes.GetNamedItem("Отчество");
                                if (at1 != null && at2 != null && at3 != null)
                                {

                                    name = ("ИНДИВИДУАЛЬНЫЙ ПРЕДПРИНИМАТЕЛЬ" + " " + "'" + "'" + at1.Value + " " + at2.Value + " " + at3.Value + "'" + "'");
                                    nameS = ("ИП" + " " + "'" + "'" + at1.Value + " " + at2.Value[0] + "." + at3.Value[0] + "." + "'" + "'");
                                }
                            }
                        }

                        foreach (XmlNode childnode in xnode.ChildNodes)
                        {
                            // получаем атрибуты с тега ОргВклМСП корневого узла
                            if (childnode.Name == "ОргВклМСП")
                            {

                                XmlNode attr2 = childnode.Attributes.GetNamedItem("НаимОрг");
                                {
                                    if (attr2 != null)
                                        name = attr2.Value;
                                }
                                XmlNode attr3 = childnode.Attributes.GetNamedItem("НаимОргСокр");
                                {
                                    if (attr3 != null)
                                        nameS = attr3.Value;
                                }
                                XmlNode attr4 = childnode.Attributes.GetNamedItem("ИННЮЛ");
                                if (attr4 != null)
                                    inn = attr4.Value;

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
                            SqlCommand cmd = new SqlCommand(@"INSERT INTO dbo.RMSP(nameOrg, nameOrgS,dateMSP,inn,categoryMSP) Values (@attr2,@attr3,@attr,@attr4,@attr1)", Connect.cnn);

                            cmd.Parameters.AddWithValue("@attr", datemsp);
                            cmd.Parameters.AddWithValue("@attr1", category);
                            cmd.Parameters.AddWithValue("@attr2", name);
                            cmd.Parameters.AddWithValue("@attr3", nameS);
                            cmd.Parameters.AddWithValue("@attr4", inn);
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
            // mainConnection.Close();
        }
    }
}
