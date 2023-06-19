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
                Connect.cnn.Open();

                foreach (XmlNode xnode in xRoot)
                {
                    DateTime datesost = DateTime.Now;
                    String NameORG = "";
                    String inn = "";
                    String nameNalog = "";
                    String summNedoimN = "";
                    String summPeni = "";
                    String summShtraf = "";
                    String ObchayaSummNedoim = "";

                    // получаем атрибуты с первого корневого узла
                    if (xnode.Attributes.Count > 0)
                    {
                        XmlNode attr = xnode.Attributes.GetNamedItem("ДатаСост");
                        if (attr != null)
                            datesost = DateTime.ParseExact(attr.Value, "mm.dd.yyyy", null);

                        foreach (XmlNode childnodes in xnode.ChildNodes)
                        {
                            if (childnodes.Name == "СведНП")
                            {
                                XmlNode attr1 = childnodes.Attributes.GetNamedItem("НаимОрг");
                                if (attr1 != null)
                                    NameORG = attr1.Value.ToUpper();
                                else
                                    NameORG = "НЕТ ИНФОРМАЦИИ";
                                XmlNode attr2 = childnodes.Attributes.GetNamedItem("ИННЮЛ");
                                if (attr2 != null)
                                    inn = attr2.Value.ToUpper();

                                foreach (XmlNode childnodess in xnode.ChildNodes)

                                {
                                    XmlNode attr3 = childnodess.Attributes.GetNamedItem("НаимНалог");
                                    if (attr3 != null)
                                        nameNalog = attr3.Value.ToUpper();

                                    XmlNode attr4 = childnodess.Attributes.GetNamedItem("СумНедНалог");
                                    if (attr4 != null)
                                        summNedoimN = attr4.Value.ToUpper();

                                    XmlNode attr5 = childnodess.Attributes.GetNamedItem("СумПени");
                                    if (attr5 != null)
                                        summPeni = attr5.Value.ToUpper();

                                    XmlNode attr6 = childnodess.Attributes.GetNamedItem("СумШтраф");
                                    if (attr6 != null)
                                        summShtraf = attr6.Value.ToUpper();

                                    XmlNode attr7 = childnodess.Attributes.GetNamedItem("ОбщСумНедоим");
                                    if (attr7 != null)
                                        ObchayaSummNedoim = attr7.Value.ToUpper();
                                }

                                int m = 0;
                                SqlCommand SqlProv = new SqlCommand(@"SELECT COUNT(ИНН) As CountTabNum FROM dbo.SpiskiRIC WHERE [ИНН]= '" + inn + "' ", Connect.cnn);
                                m = (Int32)(SqlProv.ExecuteScalar());
                            
                                if (m != 0)
                                {
                                    // внесение данных в БД
                                    SqlCommand cmd = new SqlCommand(@"INSERT INTO dbo.Nedoimki(nameORG, signDate,inn,nameNalog,summNedoimN,summPeni,summShtraf,ObchayaSummNedoim) Values (@attr1,@attr,@attr2,@attr3,@attr4,@attr5,@attr6,@attr7)", Connect.cnn);

                                    cmd.Parameters.AddWithValue("@attr", datesost);
                                    cmd.Parameters.AddWithValue("@attr1", NameORG);
                                    cmd.Parameters.AddWithValue("@attr2", inn);
                                    cmd.Parameters.AddWithValue("@attr3", nameNalog);
                                    cmd.Parameters.AddWithValue("@attr4", summNedoimN);
                                    cmd.Parameters.AddWithValue("@attr5", summPeni);
                                    cmd.Parameters.AddWithValue("@attr6", summShtraf);
                                    cmd.Parameters.AddWithValue("@attr7", ObchayaSummNedoim);

                                    cmd.ExecuteNonQuery();

                                    //добавляем новую запись в таблицу
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
            finally
            {
                Connect.cnn.Close();
            }
        }
    }
}
