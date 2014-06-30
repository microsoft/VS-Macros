using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VSMacros
{
    interface IRecorder
    {
        void StartRecording();
        void StopRecording();
    }
}
