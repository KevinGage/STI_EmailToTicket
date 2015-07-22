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


    }
}
