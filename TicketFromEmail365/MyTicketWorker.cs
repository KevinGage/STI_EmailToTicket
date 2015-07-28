using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data.SqlClient;

namespace TicketFromEmail365
{
    class MyTicketWorker
    {

        public MyTicketWorker()
        {

        }

        public static bool TestDatabaseConnection(MyConfig currentConfig)
        {
            //This should take in a config and test the connection to the database
            SqlConnectionStringBuilder builder = new System.Data.SqlClient.SqlConnectionStringBuilder();
            builder["Data Source"] = currentConfig.DbServer + "," + currentConfig.DbPort;
            builder["integrated Security"] = false;
            builder.UserID = currentConfig.UserDb;
            builder["Password"] = currentConfig.PasswordDb;
            builder["Initial Catalog"] = currentConfig.DbName;

            using (SqlConnection connection = new SqlConnection(builder.ConnectionString))
            {
                try
                {
                    connection.Open();
                    return true;
                }
                catch (SqlException)
                {
                    return false;
                }
            }

        }
    }
}
