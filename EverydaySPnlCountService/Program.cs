using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;

namespace EverydaySPnlCountService
{
    static class Program
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        static void Main()
        {
            if (Environment.UserName == "SYSTEM")
            {
                ServiceBase[] ServicesToRun;
                ServicesToRun = new ServiceBase[] { new MainProgram() };
                ServiceBase.Run(ServicesToRun);
            }
            else
            {
                MainMethod mm = new MainMethod();
                mm.Start();
                Console.ReadLine();
            }
        }
    }
}
