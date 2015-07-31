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
        string[] _ticketEmailAddresses;
        
        public MyTicketWorker(EmailMessage message, MyConfig config)
        {
            _message = message;
            _config = config;
             SubjectHasTicketNumber();
        }

        private void SubjectHasTicketNumber()
        {
            //this will fire when a new ticket worker is created.
            //it should check _message subject for exactly one instance of *** Ticket # ***
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
                        cmd.Parameters.AddWithValue("@notes", "Email from: " + _message.Sender.ToString() + System.Environment.NewLine + _message.TextBody.ToString());
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
                    ////////////////////GET ALL RELEVANT EMAIL ADDRESSES HERE AND SET _ticketEmailAddresses ALSO ADD STI GROUP
                    using (SqlCommand cmd1 = new SqlCommand("SELECT EmailAddresses FROM Tickets WHERE TicketNum=@ticketNum", conn))
                    {
                        cmd1.Parameters.AddWithValue("@ticketNumber", _ticketNumber);
                        SqlDataReader reader = cmd1.ExecuteReader();

                        if (reader.HasRows)
                        {

                        }
                        else
                        {

                        }
                    }
                    ////////////////////THIS SECTION UNFINISHED!!!!!///////////////////////////////////////

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
            string senderDomain = _message.Sender.Address.Split('@')[1];

            string connectionString = BuildConnectionString(_config);
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand("SELECT ClientEmailDomains, ClientPrimaryTech, ClientID from Clients WHERE ClientEmailDomains IS NOT NULL", conn);
                    SqlDataReader reader = cmd.ExecuteReader();

                    if(reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            string[] allDomains = reader["ClientEmailDomains"].ToString().Split(',');
                            foreach (string s in allDomains)
                            {
                                if(s == senderDomain)
                                {
                                    _primaryTech = Int32.Parse(reader["ClientPrimaryTech"].ToString());
                                    _clientID = Int32.Parse(reader["ClientID"].ToString());
                                    if (OpenTicket())
                                    {
                                        _error = null;
                                        return true;
                                    }
                                }
                            }
                        }
                        if (_config.LogLevel > 1)
                        {
                            MyLogger.writeSingleLine("email domain not found in database. Domain: " + senderDomain);
                        }
                        _error = null;
                        return false;
                    }
                    else
                    {

                        MyLogger.writeSingleLine("No email domains found in database Clients table");
                        _error = "No email domains found in database Clients table";
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                MyLogger.writeSingleLine("Error check client domain.  Error: " + ex.ToString());
                _error = "Error check client domain.  Error: " + ex.ToString();
                return false;
            }     
        }

        private bool OpenTicket()
        {
            //this should insert a new ticket using _cliendID, _primaryTech, _message.textbody as ticket issue, From address, cc addresses
            //set _ticketEmailAddresses property
            //return true if everything worked
            string connectionString = BuildConnectionString(_config);
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    //first get tech name with a select. save as variable.
                    string techName = "";

                    conn.Open();

                    using (SqlCommand cmd = new SqlCommand("SELECT TechName from Techs WHERE TechID=@techID", conn))
                    {
                        cmd.Parameters.AddWithValue("@techID", _primaryTech);

                        SqlDataReader reader = cmd.ExecuteReader();

                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                techName = reader["TechName"].ToString();
                            }

                        }
                        else
                        {
                            _error = "No tech name found";
                            return false;
                        }

                        reader.Close();
                    }
                    //then run insert.
                    if (techName != "")
                    {
                        using (SqlCommand cmd = new SqlCommand("INSERT INTO Tickets (OpenTech, AssignTech, Status, ClientID, OpenDate, Description, Priority, HoursWorked, SpecialProject, EmailAddresses) output INSERTED.TicketNum VALUES(@OpenTech, @AssignTech, @Status, @ClientID, @OpenDate, @Description, @Priority, @HoursWorked, @SpecialProject, @EmailAddresses)", conn))
                        {
                            string emailAddresses = _message.Sender.Address;
                            foreach (EmailAddress ea in _message.CcRecipients)
                            {
                                emailAddresses += "," + ea.Address;
                            }
                            foreach (EmailAddress ea in _message.BccRecipients)
                            {
                                emailAddresses += "," + ea.Address;
                            }
                            _ticketEmailAddresses = emailAddresses.Split(',');

                            cmd.Parameters.AddWithValue("@OpenTech", "EMAIL");
                            cmd.Parameters.AddWithValue("@AssignTech", techName);
                            cmd.Parameters.AddWithValue("@Status", 1);
                            cmd.Parameters.AddWithValue("@ClientID", _clientID);
                            cmd.Parameters.AddWithValue("@OpenDate", System.DateTime.Now);
                            cmd.Parameters.AddWithValue("@Description", _message.TextBody.ToString());
                            cmd.Parameters.AddWithValue("@Priority", 2);
                            cmd.Parameters.AddWithValue("@HoursWorked", 0);
                            cmd.Parameters.AddWithValue("@SpecialProject", 0);
                            cmd.Parameters.AddWithValue("@EmailAddresses", emailAddresses);

                            var newTicketNumberObject = cmd.ExecuteScalar();
                            int newTicketNumber = Int32.Parse(newTicketNumberObject.ToString());
                            _ticketNumber = newTicketNumber;

                            if (_ticketNumber > 0)
                            {
                                if (_config.LogLevel > 1)
                                {
                                    MyLogger.writeSingleLine("Ticket added to database");
                                }
                                _error = null;
                                return true;
                            }
                            else
                            {
                                _error = "Invalid ticket number returned when new ticket was created";
                                MyLogger.writeSingleLine("Invalid ticket number returned when new ticket was created");
                                return false;
                            }
                        }
                    }
                    else
                    {
                        _error = "No tech name found.";
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                _error = "Error connecting to the database.  Error: " + ex.ToString();
                MyLogger.writeSingleLine("Error connecting to the database.  Error: " + ex.ToString());
                return false;
            }
            
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
        public string[] TicketEmailAddresses
        {
            get { return _ticketEmailAddresses; }
        }
    }
}
