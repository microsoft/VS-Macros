using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VSMacros.ExecutionEngine.Stubs
{
    public class Script
    {
        public static string CreateScriptStub()
        {
            //return "dte.ExecuteCommand('File.NewFile');";

            //var script = "dte.Commands.Raise(\"{5EFC7975-14BC-11CF-9B2B-00AA00573819}\", 221, \"General\\\\Text File\", null);";

            //var script = "dte.Commands.Raise("{1496a755-94de-11d0-8c3f-00c04fc2aae2}", 13, null, null)";
            //var script = "var activeDocument = dte.ActiveDocument;";
            //var script = "dte.ItemOperations.NewFile(\"General\\\\Text File\");";

            //script += "var itemOp = dte.ItemOperations;";
            //script += "var parent = itemOp.Parent;";
            //script += "parent.ActiveDocument.NewWindow();";

            //var script = "dte.Commands.Raise(\"5efc7975-14bc-11cf-9b2b-00aa00573819\", 221, null, null);";
            //script += "dte.Commands.Raise(\"1496a755-94de-11d0-8c3f-00c04fc2aae2\", 3, null, null);";
            var script = "";
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
