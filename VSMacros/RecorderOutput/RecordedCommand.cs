//-----------------------------------------------------------------------
// <copyright file="RecordedCommand.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

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

        internal RecordedCommand() { }

        internal override void ConvertToJavascript(StreamWriter outputStream)
        {
            this.ConvertToJavascript(outputStream, 1, false);
        }

        internal void ConvertToJavascript(StreamWriter outputStream, int iterations, bool inDocument)
        {
            string output;

            if (inDocument)
            {
                outputStream.WriteLine(this.Convert(commandName, iterations));
            }
            else if (this.commandName != null)
            {
                outputStream.WriteLine(this.DuplicateStrings("dte.ExecuteCommand(\"" + this.commandName + "\");", iterations));
            }
            else
            {
                string formatString = "dte.Commands.Raise(\"{0}\", {1}{2});";
                output = string.Format(formatString, "{" + this.commandSetGuid + "}", this.commandId, (this.input == 0 ? ", null, null" : ", '" + this.input.ToString() + "', null"));
                outputStream.WriteLine(this.DuplicateStrings(output, iterations));
            }
        }

        internal void ConvertToJavascript(StreamWriter outputStream, List<char> input)
        {
           string escapedInput = string.Join("", input).Replace("\\", "\\\\").Replace("\"", "\\\"");
           string output = string.Format("Macro.InsertText(\"{0}\");", escapedInput);

            outputStream.WriteLine(output);
        }

        #region Converter

        private string Convert(string command, int iterations)
        {
            // Handled commands: "Edit.Char*", "Edit.Word*", "Edit.Line*", "Edit.BreakLine", "Edit.EndOf*", "Edit.Delete*" \ {*Column}
            //                   "Edit.Copy", "Edit.Paste", "Edit.Cut"

            // Default is dte.ExecuteCommand
            string ret = string.Empty;

            int length = command.Length;
            bool extend;

            string iterationsAsString = iterations > 1 ? iterations.ToString() : string.Empty;

            if (command.Contains("Edit."))
            {
                if (command.Contains(".Char"))
                {
                    // Edit.Char* or Edit.Char*Extend
                    extend = command.Contains("Extend");
                    string direction = command.Contains("Left") ? "Left" : "Right";

                    ret = this.FormatEditWithExtend("Char", direction, extend, iterations);
                }
                else if (command.Contains(".Word"))
                {
                    if (!command.Contains("Delete"))
                    {
                        // Edit.Word* or Edit.Word*Extend
                        extend = command.Contains("Extend");
                        string direction = command.Contains("Previous") ? "Left" : "Right";

                        ret = this.FormatEditWithExtend("Word", direction, extend, iterations);
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

                        ret = this.FormatEditWithExtend("Line", direction, extend, iterations);
                    }
                    else if (command.Contains("Start"))
                    {
                        ret = this.textSelection + "StartOfLine(1" + (extend ? ", " + extend.ToString().ToLower() : string.Empty) + ")";
                    }
                    else if (command.Contains("End"))
                    {
                        ret = this.textSelection + "EndOfLine(" + (extend ? extend.ToString().ToLower() : string.Empty) + ")";
                    }
                }
                else if (command.Contains(".Page"))
                {
                    // Edit.Page* or Edit.PAge*Extend
                    extend = command.Contains("Extend");
                    string direction = command.Contains("Up") ? "Up" : "Down";

                    ret = this.FormatEditWithExtend("Page", direction, extend, iterations);
                }
                else if (command.Contains(".Make"))
                {
                    int newCase = command.Contains("Lower") ? 1 : 2;
                    ret = this.textSelection + "ChangeCase(" + newCase + ")";
                }
                else if (command.Contains(".BreakLine"))
                {
                    ret = this.textSelection + "NewLine(" + iterationsAsString + ")";
                }
                else if (command.Contains(".Delete"))
                {
                    ret = this.textSelection + "Delete" + (command.Contains("Backwards") ? "Left" : "") + "(" + iterationsAsString + ")";
                }
                else if (command.Contains("Copy") || command.Contains("Cut") || command.Contains("Paste"))
                {
                    ret = DuplicateStrings(this.textSelection + command.Substring("Edit.".Length) + "()", iterations);
                }
                else if (command.Contains("InsertTab"))
                {
                    ret = this.textSelection + "Indent()";
                }
            }

            if (ret == string.Empty)
            {
                ret = this.DuplicateStrings("dte.ExecuteCommand(\"" + command + "\");", iterations, true);
            }

            return ret + ";";
        }

        private string DuplicateStrings(string str, int iterations, bool removeExtraSemiColon = false)
        {
            string ret = string.Concat(Enumerable.Repeat(str + Environment.NewLine, iterations));

            ret = ret.Substring(0, ret.Length - 2);

            // Remove the extra semicolon
            if (removeExtraSemiColon)
            {
                ret = ret.Substring(0, ret.Length - 1);
            }

            return ret;
        }

        private string FormatEditWithExtend(string commandName, string direction, bool extend, int iterations)
        {
            string extendLower = extend.ToString().ToLower();

            // Default for extend is false, default for count is 1
            string extendAndIterations = iterations > 1 ? extendLower + ", " + iterations.ToString() : (extend ? extendLower : string.Empty);
            return this.textSelection + commandName + direction + "(" + extendAndIterations + ")";
        }

        private string textSelection
        {
            get
            {
                return "dte.ActiveDocument.Selection.";
            }
        }

        #endregion

        internal bool IsInsert()
        {
            return this.commandName == "keyboard";
        }

        internal bool IsValidCharacter()
        {
            return this.input != '\0';
        }

        internal char Input
        {
            get { return this.input; }
        }

        internal string CommandName
        {
            get { return this.commandName; }
        }
    }
}
