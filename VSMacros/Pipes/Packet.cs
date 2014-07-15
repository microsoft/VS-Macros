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
    }
}
