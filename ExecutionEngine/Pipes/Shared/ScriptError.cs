using System;

namespace VisualStudio.Macros.ExecutionEngine.Pipes
{
    [Serializable]
    public class ScriptError
    {
        public int LineNumber;
        public int Column;
        public string Source;
        public string Description;
    } 
}
