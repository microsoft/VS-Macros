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
