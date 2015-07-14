using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.ServiceProcess;

namespace TicketFromEmail365
{
    public class TicketFromEmail365Service : ServiceBase
    {
        public TicketFromEmail365Service()
        {
            this.ServiceName = "TicketFromEmail365";
            this.CanStop = true;
            this.CanPauseAndContinue = false;
            //this.AutoLog = true;
        }

        protected override void OnStart(string[] args)
        {
           // TODO: add startup stuff
        }

        protected override void OnStop()
        {
           // TODO: add shutdown stuff
        }
    }
}
