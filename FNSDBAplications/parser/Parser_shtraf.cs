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
                Connect.cnn.Open();

                foreach (XmlNode xnode in xRoot)
                {
                    DateTime datesost = DateTime.Now;
                    String NameORG = "";
                    String inn = "";
                    String summ = "";


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
                                    XmlNode attr3 = childnodess.Attributes.GetNamedItem("СумШтраф");
                                    if (attr3 != null)
                                        summ = attr3.Value.ToUpper();

                                }
                                int m = 0;
                                SqlCommand SqlProv = new SqlCommand(@"SELECT COUNT(ИНН) As CountTabNum FROM dbo.SpiskiRIC WHERE [ИНН]= '" + inn + "' ", Connect.cnn);
                                m = (Int32)(SqlProv.ExecuteScalar());

                                if (m != 0)
                                {
                                    // внесение данных в БД
                                    SqlCommand cmd = new SqlCommand(@"INSERT INTO dbo.Pravonarush(nameORG, signDate,inn,summ) Values (@attr1,@attr,@attr2,@attr3)", Connect.cnn);

                                    cmd.Parameters.AddWithValue("@attr", datesost);
                                    cmd.Parameters.AddWithValue("@attr1", NameORG);
                                    cmd.Parameters.AddWithValue("@attr2", inn);
                                    cmd.Parameters.AddWithValue("@attr3", summ);

                                    cmd.ExecuteNonQuery();

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
