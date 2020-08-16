using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using NetOffice.Tools;
using NetOffice.OutlookApi;
using NetOffice.OutlookApi.Tools;

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
            this.OnStartupComplete += MyAddin_OnStartupComplete;
            this.OnDisconnection += MyAddin_OnDisconnection;
        }

        private void MyAddin_OnStartupComplete(ref Array custom)
        {
            Debug.WriteLine($"Running in Microsoft Outlook version {this.Application.Version}");
        }

        /// <summary>
        /// This method might not be called in Microsoft Office 2010 because of Fast Shutdown feature.
        /// </summary>
        /// <remarks>
        /// Shutdown Changes for Outlook 2010: https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ee720183(v=office.14)
        /// </remarks>
        private void MyAddin_OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            // Best Practices for Add-in Shutdown for Developers:
            // #1: Continue to release references in OnDisconnection
            // #2: Detecting when Outlook is shutting down
            // #3: Test shutdown by using both the fast and slow methods

            // While Outlook 2010 and newer no longer calls the OnBeginShutdown and OnDisconnection methods
            // of the IDTExtensibility2 interface, add-ins should still implement these methods or MyAddin_OnDisconnection
            // to release references and allocated memory, because there are scenarios in which administrators
            // might revert to slow shutdown through Group Policy, or the user might manually disconnect your
            // add-in through the COM Add-ins dialog box.
            Debug.WriteLine($"Microsoft Outlook shutdown using slow shutdown process. Disconnect mode: {removeMode}");
        }
    }
}
