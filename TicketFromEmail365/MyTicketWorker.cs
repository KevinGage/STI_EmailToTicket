using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data.SqlClient;
using Microsoft.Exchange.WebServices.Data;
using System.Text.RegularExpressions;

namespace TicketFromEmail365
{
    class MyTicketWorker
    {
        EmailMessage _message;
        MyConfig _config;
        string _error;
        int _ticketNumber;
        int _clientID;
        int _primaryTech;
        
        public MyTicketWorker(EmailMessage message, MyConfig config)
        {
            _message = message;
             SubjectHasTicketNumber();
        }

        private void SubjectHasTicketNumber()
        {
            //this will fire when a new ticket worker is created.
            //it should check _message subject for exactly one instance of *** Ticket 
            //if ticket number found set _ticketNumber and update ticket
            //if not set ticket number to 0
            MatchCollection matches = Regex.Matches(_message.Subject, "\\*{3} Ticket \\d{1,} \\*{3}");
            switch (matches.Count)
            {
                case 0:
                    _ticketNumber = 0;
                    break;
                case 1:
                    try
                    {
                        string ticketNumber = _message.Subject.Substring(_message.Subject.IndexOf(matches[0].ToString())).Split(' ')[2];
                        _ticketNumber = Int32.Parse(ticketNumber);
                        UpdateTicket();
                    }
                    catch
                    {
                        _ticketNumber = 0;
                        _error = "invalid ticket number: ";
                    }
                    break;
                default:
                    _ticketNumber = 0;
                    _error = "invalid ticket information found in subject.  Multiple tickets?";
                    break;
            }              
        }

        private void UpdateTicket()
        {
            //this should update the _ticketNumber ticket with _message in notes
            //Set error if something went wrong
            string connectionString = BuildConnectionString(_config);
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("UPDATE Tickets SET Notes = @notes + char(13)+char(10) + Notes WHERE TicketNum=@ticketNumber", conn))
                    {
                        cmd.Parameters.AddWithValue("@notes", _message.TextBody);
                        cmd.Parameters.AddWithValue("@ticketNumber", _ticketNumber);

                        int rows = cmd.ExecuteNonQuery();

                        switch (rows)
                        {
                            case 0:
                                MyLogger.writeSingleLine("No tickets updated.  Invalid ticket number? Ticket Number: " + _ticketNumber.ToString());
                                _error = "No tickets updated.  Invalid ticket number? Ticket Number: " + _ticketNumber.ToString();
                                break;
                            case 1:
                                _error = null;
                                break;
                            default:
                                MyLogger.writeSingleLine("Error updating ticket.  Multiple tickets updated? Ticket Number: " + _ticketNumber.ToString());
                                _error = "Error updating ticket.  Multiple tickets updated? Ticket Number: " + _ticketNumber.ToString();
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _error = "Error trying to update database: " + ex.ToString();
                MyLogger.writeSingleLine("Error trying to update database: " + System.Environment.NewLine + ex.ToString());
            }
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
            //this should insert a new ticket using _cliendID, _primaryTech, _message.textbody as ticket issue, From address, cc addresses
            //return true if everything worked
            return false;
        }

        public static bool TestDatabaseConnection(MyConfig currentConfig)
        {
            //This should take in a config and test the connection to the database
            string connectionString = BuildConnectionString(currentConfig);

            using (SqlConnection connection = new SqlConnection(connectionString))
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

        private static string BuildConnectionString(MyConfig config)
        {
            //This should safely build the appropriate SQL connection string using the config file
            SqlConnectionStringBuilder builder = new System.Data.SqlClient.SqlConnectionStringBuilder();
            builder["Data Source"] = config.DbServer + "," + config.DbPort;
            builder["integrated Security"] = false;
            builder.UserID = config.UserDb;
            builder["Password"] = config.PasswordDb;
            builder["Initial Catalog"] = config.DbName;

            return builder.ConnectionString;
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
