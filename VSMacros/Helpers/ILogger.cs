using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VSMacros.Helpers
{
    public interface ILogger : IDisposable
    {
        /// <summary>
        /// Will log an error messsage prepending  timestamp : ERROR: to the message
        /// </summary>
        void LogError(string component, string message, params object[] messageArgs);

        /// <summary>
        /// Will log an error messsage prepending  timestamp : INFO: to the message
        /// </summary>
        void LogInfo(string component, string message, params object[] messageArgs);

        /// <summary>
        /// Will log an error messsage prepending  timestamp : WARNING: to the message
        /// </summary>
        void LogWarning(string component, string message, params object[] messageArgs);
    }
}
