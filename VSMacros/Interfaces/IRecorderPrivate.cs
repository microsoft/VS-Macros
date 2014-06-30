using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VSMacros
{
    interface IRecorderPrivate
    {
        bool Recording { get; }
        void AddCommandData(Guid commandSet, uint identifier, string commandName, char input);
        void AddWindowActivation(Guid toolWindowID, string commandName);
        void AddWindowActivation(string path);
        void ClearData();
    }
}
