using System.ComponentModel;
using System.ServiceProcess;

namespace RunSimioModel
{
    [RunInstaller(true)]
    public partial class ProjectInstaller2 : System.Configuration.Install.Installer
    {
        private ServiceProcessInstaller process;
        private ServiceInstaller service;
        public ProjectInstaller2()
        {
            process = new ServiceProcessInstaller();
            process.Account = ServiceAccount.LocalSystem;
            service = new ServiceInstaller();
            service.ServiceName = "RunSimioModel";
            service.Description = "This Service Will Run Schedule, Save Simio Scheduled Based On Simio Project file dropped into a folder";
            Installers.Add(process);
            Installers.Add(service);
        }
    }
}
