using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VSMacros.Pipes.Packets
{
    [Serializable]
    class ScriptError
    {
        public int LineNumber;
        public int Column;
        public string Source;
        public string Description;
    }
}
