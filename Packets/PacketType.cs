using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Packets
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
        ScriptError = 3,

        /// <summary>
        /// Notifies subscribers of internal VS error.
        /// </summary>
        CriticalError = 4
    }
}
