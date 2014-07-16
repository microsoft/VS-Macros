using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VSMacros.Helpers;

namespace VSMacros.Helpers
{
    public class Logger : ILogger
    {
        public void LogError(string component, string message, params object[] messageArgs)
        {
            string error = string.Format("Error - {0}: {1}", component, message);
            Debug.WriteLine(error);
        }

        public void LogInfo(string component, string message, params object[] messageArgs)
        {
            string error = string.Format("Info - {0}: {1}", component, message);
            Debug.WriteLine(error);
        }

        public void LogWarning(string component, string message, params object[] messageArgs)
        {
            string error = string.Format("Warning - {0}: {1}", component, message);
            Debug.WriteLine(error);
        }

        public void Dispose()
        {
            
        }
    }
}
