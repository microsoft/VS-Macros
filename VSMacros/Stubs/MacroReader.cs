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

        public StreamReader CreateMacroStreamReader()
        {
            string path = Path.ChangeExtension(this.name, "txt");
            using (StreamWriter sw = new StreamWriter(path))
            {
                sw.WriteLine("dte.ExecuteCommand('File.NewFile');");
            }

            return new StreamReader(path);
        }
    }
}
