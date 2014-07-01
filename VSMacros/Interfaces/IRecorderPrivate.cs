using System;

namespace VSMacros.Interfaces
{
    interface IRecorderPrivate
    {
        bool Recording { get; }
        void AddCommandData(Guid commandSet, uint identifier, string commandName, char input);
        void AddWindowActivation(Guid toolWindowID, string windowName);
        void AddWindowActivation(string documentPath);
        void ClearData();
    }
}
