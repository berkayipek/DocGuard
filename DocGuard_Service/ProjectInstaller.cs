using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace DocGuard_Service
{
        [RunInstaller(true)]
        public class ProjectInstaller : Installer
        {
            private ServiceProcessInstaller serviceProcessInstaller;
            private ServiceInstaller serviceInstaller;

            public ProjectInstaller()
            {
                serviceProcessInstaller = new ServiceProcessInstaller();
                serviceInstaller = new ServiceInstaller();
                // Here you can set properties on serviceProcessInstaller
                //or register event handlers
                serviceProcessInstaller.Account = ServiceAccount.LocalSystem;
                serviceInstaller.ServiceName = DocGuardService.MyServiceName;
                this.Installers.AddRange(new Installer[] { serviceProcessInstaller, serviceInstaller });
            }
        }
}
