using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.IO.Compression;
using System.Diagnostics;

namespace Utils
{
    public class EventLog
    {
        public EventLog(string _Path)
        {
            ErrorServices.singleton.Proceso = "NoProceso";
            ErrorServices.singleton.Path = _Path;
        }
        public static string LogMessage(string p)
        {
            ErrorServices.singleton.LogMessage(p);
            Console.WriteLine();
            Console.WriteLine(p);
            return p;
        }

        public static string LogMessage(params object[] param)
        {
            string p = ErrorServices.singleton.LogMessage(param);
            Console.WriteLine();
            Console.WriteLine(p);
            return p;
        }

        public static void LogError(ErrorType errorType, params object[] param)
        {
            LogMessage(param);
            string p = ErrorServices.singleton.LogError(errorType, null, param);
            Console.WriteLine();
            Console.WriteLine(p);
        }
        public static void LogError(ErrorType errorType, Exception ex, params object[] param)
        {

            ErrorServices.singleton.LogError(errorType, ex, param);
            string p = DateTime.Now.ToShortTimeString() + " ->" + errorType.ToString() + ": " + ex.Message + "\r\n" + ex.StackTrace;

            if (ex.InnerException != null)
            {
                Console.WriteLine(ex.InnerException.Message);
                p += "\r\n->" + ex.InnerException.Message;
            }
            ErrorServices.singleton.LogMessage("Error:", p, param);


            Console.WriteLine();
            Console.WriteLine(p);
        }

        public string Proceso {
            set
            {
                ErrorServices.singleton.Proceso = value; 
            }
        }
    }
 
  
}