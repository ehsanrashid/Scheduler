using System;
using System.Collections.Generic;
using System.Text;
using System.ServiceProcess;

namespace SchedulerService
{
    static class Startup
    {
        // The main entry point for the process
        [STAThread]
        public static void Main()
        {
            // More than one user Service may run within the same process.
            var ServicesToRun = new ServiceBase[] 
                { 
                    new ScheduleService(),

                };

            try
            {
                ServiceBase.Run(ServicesToRun);
            }
            catch (Exception)
            { }
        }
    }
}
