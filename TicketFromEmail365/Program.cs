﻿using System;
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
            System.ServiceProcess.ServiceBase.Run(new TicketFromEmail365Service());
        }
    }
}
