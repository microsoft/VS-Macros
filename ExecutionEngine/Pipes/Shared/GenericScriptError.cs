using System;

namespace VisualStudio.Macros.ExecutionEngine.Pipes
{
    [Serializable]
    public class GenericScriptError
    {
        public int LineNumber;
        public int Column;
        public string Source;
        public string Description;
    } 
}
