//-----------------------------------------------------------------------
// <copyright file="InternalVSException.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;

namespace VSMacros.ExecutionEngine
{
    public class InternalVSException : Exception
    {
        public InternalVSException(string message, string source, string stackTrace, string targetSite)
        {
            this.Description = message;
            this.Source = source;
            this.StackTrace = stackTrace;
            this.TargetSite = targetSite;
        }

        public string Description { get; private set; }
        public override string Source { get; set; }
        public new string StackTrace { get; private set; }
        public new string TargetSite { get; private set; }
    }
}
