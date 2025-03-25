using System.ComponentModel;
using System.Configuration.Install;
using System.ServiceProcess;

namespace RunSimioModelObjects
{
    [RunInstaller(true)]
    public partial class ProjectInstaller : Installer
    {
        private ServiceProcessInstaller process;
        private ServiceInstaller service;
        public ProjectInstaller()
        {
            process = new ServiceProcessInstaller();
            process.Account = ServiceAccount.LocalSystem;
            service = new ServiceInstaller();
            service.ServiceName = "RunSimioModelObjects";
            service.Description = "This Service Will Import Data, Run Schedule, Export Schedule and Save Simio Scheduled Based On The Creation Of A File";
            Installers.Add(process);
            Installers.Add(service);
        }
    }
}
