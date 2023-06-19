using System;
using System.Data.SqlClient;

namespace FNSDBAplications.connection
{
    public class Connect : IDisposable
    {
        public static SqlConnection cnn = new SqlConnection(@"Data Source=DPO-STAT1\SQLEXPRESS;Initial Catalog=fns;Integrated Security=True");

        public void Dispose()
        {
            if (cnn != null)
            {
                cnn.Dispose();
            }
        }

    }
}
