using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MicrosoftCorporation.VSMacros.Pipes.Packets
{
    [Serializable]
    public class CriticalError
    {
        public string Message;
        public string Source;
        public string StackTrace;
        public string TargetSite;
    }
}
