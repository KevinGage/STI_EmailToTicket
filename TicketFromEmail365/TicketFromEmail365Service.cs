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
            Logger.checkLogFile();
            Logger.writeSingleLine("Program started in service mode");

            Config conf = new Config(@".\TicketsFromEmail365.cfg");

            if (conf.Error != "")
            {
                Logger.writeSingleLine(conf.Error);
                Logger.writeSingleLine("Terminating");
                // do something here to stop service
            }
            else
            {
                Logger.writeSingleLine(@"Succesfully Read Config File: .\TicketsFromEmail365.cfg");
                EwsWorker worker = new EwsWorker(conf);
            }
        }

        protected override void OnStop()
        {
           // TODO: add shutdown stuff for running as service
        }

        public void Start()
        {
            // TODO: add startup stuff for running as console for debugging
            // Delete function when done developing program
            Logger.checkLogFile();
            Logger.writeSingleLine("Program started in debugging mode");
        }

        public void DoStop()
        {
            // TODO: add shutdown stuff for running as console for debugging
            // Delete function when done developing program
        }
    }
}
