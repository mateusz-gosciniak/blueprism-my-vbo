using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace Kill_Process
{
    class Program
    {
        // INPUT
        static public string ProcessName = "chrome";
        // OUTPUT
        static public bool Success;
        static public string Error_Message;
        static public void KillProcess(string ProcessName)
        {
            var processes = Process.GetProcessesByName(ProcessName);
            foreach (var p in processes)
            {
                p.Kill();
            }
        }
        static void Main(string[] args)
        {
            Error_Message = string.Empty;
            try
            {
                KillProcess(ProcessName);
                Success = true;
            }
            catch(Exception ex)
            {
                Success = false;
                Error_Message = "Failure to Kill Process. " + ex.Message;
            }
        }
    }
}
