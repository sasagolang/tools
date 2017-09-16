using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace WindowsService1
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            // Interop.ShowMessageBox("This a message from AlertService.",
            //          "AlertService Message");
            Interop.CreateProcess("1.bat", @"C:\Debug\");
            //  ApplicationLoader.PROCESS_INFORMATION aa;
            // ApplicationLoader.StartProcessAndBypassUAC(@"C:\Debug\1.bat", out  aa);
        }

        protected override void OnStop()
        {
        }
    }
}
