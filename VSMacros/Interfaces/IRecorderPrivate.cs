//-----------------------------------------------------------------------
// <copyright file="IRecorderPrivate.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;

namespace VSMacros.Interfaces
{
    interface IRecorderPrivate
    {
        bool IsRecording { get; }
        void AddCommandData(Guid commandSet, uint identifier, string commandName, char input);
        void AddWindowActivation(Guid toolWindowID, string windowName);
        void AddWindowActivation(string documentPath);
        void ClearData();
    }
}
