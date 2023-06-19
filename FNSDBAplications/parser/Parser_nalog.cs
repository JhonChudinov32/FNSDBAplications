using System;
using System.Data.SqlClient;
using System.Windows;
using System.Xml;
using FNSDBAplications.connection;

namespace FNSDBAplications.parser
{
    public class Parser_nalog
    {
        //Парсер налогов
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
                    String summUplnalog = "";

                    if (xnode.Attributes.Count > 0)
                    {
                        XmlNode attr = xnode.Attributes.GetNamedItem("ДатаСост");
                        if (attr != null)
                            datesost = DateTime.ParseExact(attr.Value, "mm.dd.yyyy", null);

                        // получаем атрибуты с первого корневого узла
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


                                foreach (XmlNode child in xnode.SelectNodes("СвУплСумНал"))
                                {

                                    if (child.Attributes.Count > 0)
                                    {

                                        XmlNode attr3 = child.Attributes.GetNamedItem("НаимНалог");
                                        if (attr3 != null)
                                            nameNalog = attr3.Value;

                                        XmlNode attr4 = child.Attributes.GetNamedItem("СумУплНал");
                                        if (attr4 != null)
                                            summUplnalog = attr4.Value.ToUpper();

                                    }
                                    int m = 0;
                                    SqlCommand SqlProv = new SqlCommand(@"SELECT COUNT(ИНН) As CountTabNum FROM dbo.SpiskiRIC WHERE [ИНН]= '" + inn + "' ", Connect.cnn);
                                    
                                    m = (Int32)(SqlProv.ExecuteScalar());
                                 

                                    if (m != 0)
                                    {
                                        // внесение данных в БД
                                        SqlCommand cmd = new SqlCommand(@"INSERT INTO dbo.Nalog(nameORG,signDate, inn,nameNalog,summUplnalog) Values (@attr1,@attr,@attr2,@attr3,@attr4)", Connect.cnn);

                                        cmd.Parameters.AddWithValue("@attr", datesost);
                                        cmd.Parameters.AddWithValue("@attr1", NameORG);
                                        cmd.Parameters.AddWithValue("@attr2", inn);
                                        cmd.Parameters.AddWithValue("@attr3", nameNalog);
                                        cmd.Parameters.AddWithValue("@attr4", summUplnalog);

                                        cmd.ExecuteNonQuery();

                                        //добавляем новую запись в таблицу

                                        
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
            finally
            {
                Connect.cnn.Close();
            }
        }
   
    }
}
