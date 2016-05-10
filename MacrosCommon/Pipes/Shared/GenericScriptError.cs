//-----------------------------------------------------------------------
// <copyright file="ScriptError.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

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
