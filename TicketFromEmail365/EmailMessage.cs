using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TicketFromEmail365
{
    class EmailMessage
    {
        string _sender;
        string _subject;
        string _body;

        public EmailMessage(string sender, string subject, string body)
        {
            _sender = sender;
            _subject = subject;
            _body = body;
        }

        public bool CheckEmailDomain()
        {
            //should check database for senders email domain.  Return true if client domain exists. false if domain doesn't belong to a client.
            return false;
        }

        private void ProcessSubject()
        {
            //This should be triggered by CheckEmailDomain if the message is from a client.
            //First check subject to see if it already includes a ticket number.  If yes verify ticket exists and is open.  Re-open if closed.
            //If no create new ticket in database for appropriate client.  Change current subject to include ticket info
        }
    }
}
