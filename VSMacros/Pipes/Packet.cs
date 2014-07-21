//-----------------------------------------------------------------------
// <copyright file="Packet.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace VSMacros.Enums
{
    /// <summary>
    /// Specifies which type of packet the stream is receiving.
    /// </summary>
    internal enum Packet
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
        ScriptError = 3,

        /// <summary>
        /// Notifies subscribers of internal VS error.
        /// </summary>
        CriticalError = 4
    }
}
