using System;
using System.Data.SqlClient;

namespace FNSDBAplications.connection
{
    public class Connect : IDisposable
    {
        //public static SqlConnection cnn = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;Initial Catalog=fns;Integrated Security=True;Connect Timeout=30");
        public static SqlConnection cnn = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=F:\Проекты\FNSDBAplications\FNSDBAplications\bin\Debug\fns.mdf;Integrated Security=True; Connect Timeout=30");

        public void Dispose()
        {
            if (cnn != null)
            {
                cnn.Dispose();
            }
        }

    }
}
