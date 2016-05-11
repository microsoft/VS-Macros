//-----------------------------------------------------------------------
// <copyright file="PacketType.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;

namespace VisualStudio.Macros.ExecutionEngine.Pipes
{
    [Serializable]
    public enum PacketType
    {
        /// <summary>
        /// Empty file packet.
        /// </summary>
        Empty = -1,

        /// <summary>
        /// Contains information on the file path for a file containing a macro to execute.
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
        GenericScriptError = 3,

        /// <summary>
        /// Notifies subscribers of internal VS error.
        /// </summary>
        CriticalError = 4
    }
}
