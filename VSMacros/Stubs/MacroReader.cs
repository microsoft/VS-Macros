using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MicrosoftCorporation.VSMacros.Stubs
{
    class MacroReader
    {
        string name = "currentScript";

        internal string AppendExtension(string name)
        {
            return name + ".txt";
        }

        public StreamReader CreateMacroStreamReader()
        {
            var path = AppendExtension(this.name);
            using (StreamWriter sw = new StreamWriter(path))
            {
                sw.WriteLine("dte.ExecuteCommand('File.NewFile');");
            }

            return new StreamReader(path);
        }
    }
}
