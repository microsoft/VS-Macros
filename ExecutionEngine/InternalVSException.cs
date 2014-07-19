using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        public string StackTrace { get; private set; }
        public string TargetSite { get; private set; }
    }
}
