//-----------------------------------------------------------------------
// <copyright file="InputParser.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Globalization;
using System.IO;
using Microsoft.Internal.VisualStudio.Shell;
using VisualStudio.Macros.ExecutionEngine;
using VSMacros.ExecutionEngine.Pipes.Shared;

namespace ExecutionEngine.Helpers
{
    public static class InputParser
    {
        private static string[] stringSeparator = new [] { SharedVariables.Delimiter };

        internal static string[] SeparateArgs(string[] args)
        {
            string[] separatedArgs = args[0].Split(InputParser.stringSeparator, StringSplitOptions.RemoveEmptyEntries);
            return separatedArgs;
        }

        internal static int GetPid(string unparsedPid)
        {
            int pid;

            Validate.IsNotNullAndNotEmpty(unparsedPid, "unparsedPid");

            if (!int.TryParse(unparsedPid, out pid))
            {
                throw new ArgumentException(string.Format(CultureInfo.CurrentUICulture, Resources.InvalidPIDArgument, unparsedPid, "unparsedPid"));
            }

            return pid;
        }

        internal static Guid GetGuid(string unparsedGuid)
        {
            var guid = Guid.Parse(unparsedGuid);
            Validate.IsNotNullAndNotEmpty(unparsedGuid, "unparsedGuid");
            return guid;
        }

        internal static string ExtractScript(string path)
        {
            Validate.IsNotNullAndNotEmpty(path, "path");
            return File.ReadAllText(path);
        }

        internal static string WrapScript(string unwrapped)
        {
            //string s = "if (typeof(dte.ActiveDocument) != null) { dte.ActiveDocument.Activate(); }";

            string activateActiveDocument = unwrapped.Contains("ExecuteCommand(\"Edit.") ? "dte.ActiveDocument.Activate()" : string.Empty;
            //string activateActiveDocument = Regex.IsMatch(unwrapped, "//.[^$]*ExecuteCommand(\"Edit.") ? "dte.ActiveDocument.Activate()" : string.Empty;
            return string.Format("function {0}() {{{1}{3}{1}{2}{1}}}", Program.MacroName, Environment.NewLine, unwrapped, activateActiveDocument);
        }
    }
}
