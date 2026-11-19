using Common;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TesoreriaCobranza
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                if (Process.GetProcessesByName(Process.GetCurrentProcess().ProcessName).Length > 1)
                {
                    //instance of the process already active, exit
                    return;
                }
                ErrorServices.singleton.Path = UtilityIO.MacroLogPath;
                DirectoryInfo directoryInfo = new DirectoryInfo(UtilityIO.MacroTempPath);
                if (directoryInfo.Exists)
                    try { directoryInfo.Delete(true); } 
                    catch { }

                Application.EnableVisualStyles();
                //Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new frmTesoreriaCobranza());
            }
            catch (Exception ex)
            {
                Utils.EventLog.LogError(ErrorType.ERROR, ex, ex.StackTrace);
                Application.Exit();
            }
        }
    }
}
