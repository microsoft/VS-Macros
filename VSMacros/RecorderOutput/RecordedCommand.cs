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
            else if (this.commandName == "keyboard")
            {
                string formatString = "cmdHelper.DispatchCommandWithArgs(\"{0}\", {1}{2})";
                output = string.Format(formatString, "{" + this.commandSetGuid + "}", this.commandId, (this.input == 0 ? ", null" : ", \"" + this.input.ToString() + "\""));
                outputStream.WriteLine(output);
            }
            else
            {
                outputStream.WriteLine(Parse(commandName));
            }
        }

        private string Parse(string command)
        {
            // Handled commands: "Edit.Char*", "Edit.Word*", "Edit.Line*", "Edit.BreakLine", "Edit.EndOf*", "Edit.Delete*"
            //                   "Edit.Copy", "Edit.Paste", "Edit.Cut"

            // Default is dte.ExecuteCommand
            string ret = "dte.ExecuteCommand(\"" + command + "\")";

            int length = command.Length;
            bool extend;

            if (command.Contains("Edit."))
            {
                if (command.Contains(".Char"))
                {
                    // Edit.Char* or Edit.Char*Extend
                    extend = command.Contains("Extend");
                    string direction = command.Contains("Left") ? "Left" : "Right";

                    ret = this.FormatEditWithExtend("Char", direction, extend);
                }
                else if (command.Contains(".Word"))
                {
                    if (!command.Contains("Delete"))
                    {
                        // Edit.Word* or Edit.Word*Extend
                        extend = command.Contains("Extend");
                        string direction = command.Contains("Previous") ? "Left" : "Right";

                        ret = this.FormatEditWithExtend("Word", direction, extend);
                    }
                }
                else if (command.Contains(".Line"))
                {
                    // Edit.Line* or Edit.Line*Extend
                    extend = command.Contains("Extend");
                    int extendIndex = command.Length,
                        editLineLength = "Edit.Line".Length;

                    if (command.Contains("Up") || command.Contains("Down"))
                    {
                        if (extend)
                        {
                            extendIndex = command.IndexOf("Extend");
                        }

                        string direction = command.Substring(editLineLength, extendIndex - editLineLength);

                        ret = this.FormatEditWithExtend("Line", direction, extend);
                    }
                    else if (command.Contains("Start"))
                    {
                        ret = this.textSelection + "StartOfLine(0, " + extend.ToString().ToLower() + ")";
                    }
                    else if (command.Contains("End"))
                    {
                        ret = this.textSelection + "EndOfLine(" + extend.ToString().ToLower() + ")";
                    }
                }
                else if (command.Contains(".Page"))
                {
                    // Edit.Page* or Edit.PAge*Extend
                    extend = command.Contains("Extend");
                    string direction = command.Contains("Up") ? "Up" : "Down";

                    ret = this.FormatEditWithExtend("Page", direction, extend);
                }
                else if (command.Contains(".BreakLine"))
                {
                    ret = this.textSelection + "NewLine(1)";
                }
                else if (command.Contains(".EndOf")) { }
                else if (command.Contains(".Delete"))
                {
                    ret = command.Contains("Backwards") ? this.textSelection + "DeleteLeft(1)" : this.textSelection + "Delete(1)";
                }
                else if (command.Contains("Copy") || command.Contains("Cut") || command.Contains("Paste"))
                {
                    ret = this.textSelection + command.Substring("Edit.".Length) + "()";
                }

            }

            return ret + ";";
        }

        private string FormatEditWithExtend(string commandName, string direction, bool extend)
        {
            return this.textSelection + commandName + direction + "(" + extend.ToString().ToLower() + ", 1)";
        }

        private string textSelection
        {
            get
            {
                return "dte.ActiveDocument.Selection.";
            }
        }

    }
}
