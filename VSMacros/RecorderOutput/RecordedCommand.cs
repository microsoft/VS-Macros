﻿//-----------------------------------------------------------------------
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
            string formatString = "dte.Commands.Raise(\"{0}\", {1}{2})";
            string output = string.Format(formatString, this.commandSetGuid, this.commandId, (this.input == 0 ? string.Empty : ", " + this.input.ToString()));

            outputStream.WriteLine(output);
        }
    }
}