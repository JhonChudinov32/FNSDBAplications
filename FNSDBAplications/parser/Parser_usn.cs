using System;
using System.Data.SqlClient;
using System.Windows;
using System.Xml;
using FNSDBAplications.connection;


namespace FNSDBAplications.parser
{
    public class Parser_usn
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
                    String ESXN = "";
                    String USN = "";
                    String EIVD = "";
                    String SRP = "";

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
                                    NameORG = attr1.Value;

                                XmlNode attr2 = childnodes.Attributes.GetNamedItem("ИННЮЛ");
                                if (attr2 != null)
                                    inn = attr2.Value;

                            }
                            foreach (XmlNode childnodess in xnode.ChildNodes)

                            {
                                XmlNode attr3 = childnodess.Attributes.GetNamedItem("ПризнЕСХН");
                                if (attr3 != null)
                                    ESXN = attr3.Value;
                                XmlNode attr4 = childnodess.Attributes.GetNamedItem("ПризнУСН");
                                if (attr4 != null)
                                    USN = attr4.Value;
                                XmlNode attr5 = childnodess.Attributes.GetNamedItem("ПризнЕНВД");
                                if (attr5 != null)
                                    EIVD = attr5.Value;
                                XmlNode attr6 = childnodess.Attributes.GetNamedItem("ПризнСРП");
                                if (attr6 != null)
                                    SRP = attr6.Value;
                            }
                        }
                        int m = 0;
                        SqlCommand SqlProv = new SqlCommand(@"SELECT COUNT(ИНН) As CountTabNum FROM dbo.SpiskiRIC WHERE [ИНН]= '" + inn + "' ", Connect.cnn);
                        m = (Int32)(SqlProv.ExecuteScalar());

                        if (m != 0)
                        {
                            // внесение данных в БД
                            SqlCommand cmd = new SqlCommand(@"INSERT INTO dbo.USN(nameORG, signDate,inn,esxn,usn,eivd,srp) Values (@attr1,@attr,@attr2,@attr3,@attr4,@attr5,@attr6)", Connect.cnn);

                            cmd.Parameters.AddWithValue("@attr", datesost);
                            cmd.Parameters.AddWithValue("@attr1", NameORG);
                            cmd.Parameters.AddWithValue("@attr2", inn);
                            cmd.Parameters.AddWithValue("@attr3", ESXN);
                            cmd.Parameters.AddWithValue("@attr4", USN);
                            cmd.Parameters.AddWithValue("@attr5", EIVD);
                            cmd.Parameters.AddWithValue("@attr6", SRP);
                            cmd.ExecuteNonQuery();

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
