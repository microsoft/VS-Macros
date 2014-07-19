//-----------------------------------------------------------------------
// <copyright file="Packet.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExecutionEngine.Enums
{
    internal enum Packet
    {
        /// <summary>
        /// Contains information to parse a file path.
        /// </summary>
        FilePath = 0,

        /// <summary>
        /// Sends a request to close the engine.
        /// </summary>
        Close = 1,

        /// <summary>
        /// Notifies subscribers of success completion of macro.
        /// </summary>
        Success = 2,

        /// <summary>
        /// Notifies subscribers of an error upon running the script.
        /// </summary>
        ScriptError = 3,

        /// <summary>
        /// Notifies subscribers of internal VS error.
        /// </summary>
        CriticalError = 4
    }
}
