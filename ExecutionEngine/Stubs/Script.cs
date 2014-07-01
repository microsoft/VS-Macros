using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VisualStudio.Macros.ExecutionEngine.Stubs
{
    public class Script
    {
        public string CreateScriptStub()
        {
            //return "dte.ExecuteCommand('File.NewFile');";
            var script = "var activeDocument = dte.ActiveDocument;";
            script += "var itemOp = dte.ItemOperations;";
            script += "var parent = itemOp.Parent;";
            script += "parent.ActiveDocument.NewWindow();";

            return script;
        }
    }
}
