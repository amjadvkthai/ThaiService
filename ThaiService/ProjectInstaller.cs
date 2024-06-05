using System.ComponentModel;
using System.Configuration.Install;
using System.ServiceProcess;

namespace UserInformationService
{
    [RunInstaller(true)]
    public class ProjectInstaller : Installer
    {
        private ServiceProcessInstaller processInstaller;
        private ServiceInstaller serviceInstaller;

        public ProjectInstaller()
        {
            processInstaller = new ServiceProcessInstaller();
            processInstaller.Account = ServiceAccount.LocalSystem;

            serviceInstaller = new ServiceInstaller();
            serviceInstaller.StartType = ServiceStartMode.Automatic;
            serviceInstaller.ServiceName = "UserInformationService";
            serviceInstaller.Description = "Thai Monitoring Software";

            Installers.Add(processInstaller);
            Installers.Add(serviceInstaller);
        }
    }
}
