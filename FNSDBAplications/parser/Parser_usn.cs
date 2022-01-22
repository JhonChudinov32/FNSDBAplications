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

                foreach (XmlNode xnode in xRoot)
                {
                    DateTime signDate = DateTime.Now;
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
                        SqlCommand SqlProv = new SqlCommand(@"SELECT COUNT(inn) As CountTabNum FROM dbo.CompanyGroup WHERE inn= '" + inn + "' ", Connect.cnn);
                        Connect.cnn.Open();
                        m = (Int32)(SqlProv.ExecuteScalar());
                        Connect.cnn.Close();

                        if (m != 0)
                        {

                            // подключение к БД
                            Connect.cnn.Open();
                            // внесение данных в БД
                            SqlCommand cmd = new SqlCommand(@"INSERT INTO dbo.USN(signDate,inn,esxn,usn,eivd,srp) Values (@attr,@attr2,@attr3,@attr4,@attr5,@attr6)", Connect.cnn);

                            cmd.Parameters.AddWithValue("@attr", signDate);
                            cmd.Parameters.AddWithValue("@attr2", inn);
                            cmd.Parameters.AddWithValue("@attr3", ESXN);
                            cmd.Parameters.AddWithValue("@attr4", USN);
                            cmd.Parameters.AddWithValue("@attr5", EIVD);
                            cmd.Parameters.AddWithValue("@attr6", SRP);

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
