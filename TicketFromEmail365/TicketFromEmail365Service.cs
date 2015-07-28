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
            if (!MyLogger.checkLogFile())
            {
                Environment.Exit(1); //not sure if this works
            }
            MyLogger.writeSingleLine("Program started in service mode");

            MyConfig conf = new MyConfig(@".\TicketsFromEmail365.cfg");

            if (conf.Error != "")
            {
                MyLogger.writeSingleLine(conf.Error);
                MyLogger.writeSingleLine("Terminating");
                Environment.Exit(1); //not sure if this works
            }
            MyLogger.writeSingleLine(@"Succesfully Read Config File: .\TicketsFromEmail365.cfg");

            if (!MyTicketWorker.TestDatabaseConnection(conf))
            {
                MyLogger.writeSingleLine("Error testing database connection");
                MyLogger.writeSingleLine("Terminating");
                Environment.Exit(1); //not sure if this works
            }
            MyLogger.writeSingleLine(@"Succesfully tested database connection.");

            MyEwsWorker worker = new MyEwsWorker(conf);

            //Not sure if something is needed here to keep service running.  Or will the worker object keep the service alive?
            if (worker.Error == "")
            {
                //Do stuff to keep the service running???
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
            
            MyLogger.writeSingleLine("Program started in debugging mode");
        }

        public void DoStop()
        {
            // TODO: add shutdown stuff for running as console for debugging
            // Delete function when done developing program
        }
    }
}
