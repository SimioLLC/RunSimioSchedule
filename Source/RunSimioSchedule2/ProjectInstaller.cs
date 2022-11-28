using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.ServiceProcess;
using System.Threading.Tasks;

namespace RunSimioSchedule2
{
    [RunInstaller(true)]
    public partial class ProjectInstaller2 : Installer
    {
        private ServiceProcessInstaller process;
        private ServiceInstaller service;
        public ProjectInstaller2()
        {
            process = new ServiceProcessInstaller();
            process.Account = ServiceAccount.LocalSystem;
            service = new ServiceInstaller();
            service.ServiceName = "RunSimioSchedule2";
            service.Description = "This Service Will Run Schedule, Save Simio Scheduled Based On Simio Project file dropped into a folder";
            Installers.Add(process);
            Installers.Add(service);
        }
    }
}
