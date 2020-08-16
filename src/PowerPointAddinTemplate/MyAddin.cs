using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using NetOffice.Tools;
using NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Tools;

namespace $safeprojectname$
{
    [ComVisible(true)]
    [Guid("$guid3$")]
    [ProgId("$safeprojectname$.MyAddin")]
    [COMAddin("MyAddin", "Addin description.", LoadBehavior.LoadAtStartup)]
    public class MyAddin : COMAddin
    {
        public MyAddin()
        {
            this.OnConnection += MyAddin_OnConnection;
        }

        private void MyAddin_OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            this.Application.PresentationOpenEvent += Application_PresentationOpenEvent;
        }

        private void Application_PresentationOpenEvent(Presentation presentation)
        {
            using (presentation)
            {
                // start working with the workbook
            }
        }
    }
}
