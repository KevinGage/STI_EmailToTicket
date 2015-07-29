using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data.SqlClient;
using Microsoft.Exchange.WebServices.Data;

namespace TicketFromEmail365
{
    class MyTicketWorker
    {
        EmailMessage _message;
        string _error;
        int _ticketNumber;
        int _clientID;
        int _primaryTech;
        
        public MyTicketWorker(EmailMessage message)
        {
            _message = message;
            SubjectHasTicketNumber();
        }

        private void SubjectHasTicketNumber()
        {
            //this will fire when a new ticket worker is created.
            //it should check _message subject for *** Ticket # ***
            //if ticket number found set _ticketNumber and update ticket
            //if not set ticket number to 0
            _ticketNumber = 0;
        }

        public bool UpdateTicket()
        {
            //this should update the _ticketNumber ticket with _message in notes
            //return false if error
            return false;
        }

        public bool CheckClientDomain()
        {
            //this should check the _message sender domain and see if it is in the clients table.
            //if not exactly 1 result found return false
            //if 1 result found set _clientID and _primaryTech, OpenTicket().
            //if open ticket succeeds return true.
            return false;
        }

        private bool OpenTicket()
        {
            //this should insert a new ticket using _cliendID, _primaryTech, _message.textbody, openedby?
            //return true if everything worked
            return false;
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

        public int TicketNumber
        {
            get { return _ticketNumber; }
        }
        public string Error
        {
            get { return _error; }
        }
    }
}
