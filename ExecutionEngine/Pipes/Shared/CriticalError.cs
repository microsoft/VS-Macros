//-----------------------------------------------------------------------
// <copyright file="CriticalError.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;

namespace VisualStudio.Macros.ExecutionEngine.Pipes
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
