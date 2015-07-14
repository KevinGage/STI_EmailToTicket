using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.ComponentModel;
using System.Configuration.Install;
using System.ServiceProcess;


namespace TicketFromEmail365
{
    // to install this service InstallUtil /LogToConsole=true TicketFromEmail365.exe
    [RunInstaller(true)]
    public class TicketFromEmail365Installer : Installer
    {
        private ServiceProcessInstaller processInstaller;
        private ServiceInstaller serviceInstaller;

        public TicketFromEmail365Installer()
        {
            processInstaller = new ServiceProcessInstaller();
            serviceInstaller = new ServiceInstaller();

            processInstaller.Account = ServiceAccount.LocalSystem;
            serviceInstaller.StartType = ServiceStartMode.Automatic;
            serviceInstaller.ServiceName = "TicketFromEmail365"; //must match TicketFromEmail365Service.ServiceName
            serviceInstaller.Description = "Service to connect office 365 emails to STI Ticket System";

            Installers.Add(serviceInstaller);
            Installers.Add(processInstaller);
        } 
    }
}
