using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.ServiceProcess;

namespace TicketFromEmail365
{
    class Program
    {
        static void Main()
        {
            // This used to run the service as a console (development phase only)
            // When done with development uncomment line below and remove all code between this

            var serviceToRun = new TicketFromEmail365Service();

            if (Environment.UserInteractive)
            {
                serviceToRun.Start();

                Config conf = new Config(@".\TicketsFromEmail365.cfg");

                if (conf.Error != "")
                {
                    Logger.writeSingleLine(conf.Error);
                    Logger.writeSingleLine("Terminating");
                    serviceToRun.DoStop();
                }
                else
                {
                    Logger.writeSingleLine(@"Succesfully Read Config File: .\TicketsFromEmail365.cfg");
                    EwsWorker worker = new EwsWorker(conf);
                }

                Console.WriteLine("Press Enter to terminate ...");
                Console.ReadLine();

                serviceToRun.DoStop();
            }
            else
            {
                ServiceBase.Run(serviceToRun);
            }

            // When done with development uncomment line below and remove all code between this

            //System.ServiceProcess.ServiceBase.Run(new TicketFromEmail365Service());
        }
    }
}
