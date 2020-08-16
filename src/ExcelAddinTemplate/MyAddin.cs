using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using NetOffice.Tools;
using NetOffice.ExcelApi;
using NetOffice.ExcelApi.Tools;

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
            this.Application.WorkbookOpenEvent += Application_WorkbookOpenEvent;
        }

        private void Application_WorkbookOpenEvent(Workbook workbook)
        {
            using (workbook)
            {
                // start working with the workbook
            }
        }
    }
}
