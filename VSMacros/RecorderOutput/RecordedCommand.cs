//-----------------------------------------------------------------------
// <copyright file="RecordedCommand.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.IO;

namespace VSMacros.RecorderOutput
{
    [DebuggerDisplay("{commandName}")]
    internal sealed class RecordedCommand : RecordedActionBase
    {
        private readonly Guid commandSetGuid;
        private readonly uint commandId;
        private readonly string commandName;
        private readonly char input;

        internal RecordedCommand(Guid commandSetGuid, uint commandId, string commandName, char input)
        {
            this.commandName = commandName;
            this.commandSetGuid = commandSetGuid;
            this.commandId = commandId;
            this.input = input;
        }

        internal override void ConvertToJavascript(StreamWriter outputStream)
        {
            string output;
            if (this.commandName == null)
            {
                string formatString = "dte.Commands.Raise(\"{0}\", {1}{2})";
                output = string.Format(formatString, "{" + this.commandSetGuid + "}", this.commandId, (this.input == 0 ? ", null, null" : ", '" + this.input.ToString() + "', null"));
                outputStream.WriteLine(output);
            }
            else
            {
                outputStream.WriteLine("dte.ExecuteCommand(\"" + this.commandName + "\")");
            }
        }
    }
}
