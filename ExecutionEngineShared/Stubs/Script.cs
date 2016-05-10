//-----------------------------------------------------------------------
// <copyright file="Script.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace VSMacros.ExecutionEngine.Stubs
{
    public class Script
    {
        public static string CreateScriptStub()
        {
            var script = string.Empty;
            return script;
        }

        public static string OpenFileDialogStub()
        {
            return "dte.ExecuteCommand('File.NewFile');";
        }

        public static string OpenFileStub()
        {
            return "dte.Commands.Raise(\"{5EFC7975-14BC-11CF-9B2B-00AA00573819}\", 221, \"General\\\\Text File\", null);";
        }
    }
}
