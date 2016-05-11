//-----------------------------------------------------------------------
// <copyright file="RuntimeException.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;

namespace VSMacros.ExecutionEngine
{
    public class RuntimeException : Exception
    {
        public RuntimeException(string message, string source, uint lineNumber, int characterPosition)
        {
            this.Description = message;
            this.Source = source;
            this.Line = lineNumber;
            this.CharacterPosition = characterPosition;
        }

        public string Description { get; private set; }
        public uint Line { get; private set; }
        public int CharacterPosition { get; private set; }
        public override string Source { get; set; }
    }
}
