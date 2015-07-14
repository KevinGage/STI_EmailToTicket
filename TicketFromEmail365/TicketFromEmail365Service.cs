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
           // TODO: add startup stuff for running as service
        }

        protected override void OnStop()
        {
           // TODO: add shutdown stuff for running as service
        }

        public void Start()
        {
            // TODO: add startup stuff for running as console for debugging
            // Delete function when done developing program
        }

        public void DoStop()
        {
            // TODO: add shutdown stuff for running as console for debugging
            // Delete function when done developing program
        }
    }
}
