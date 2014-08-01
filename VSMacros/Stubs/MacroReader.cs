//-----------------------------------------------------------------------
// <copyright file="MacroReader.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.IO;

namespace VSMacros.Stubs
{
    class MacroReader
    {
        string name = "currentScript";

        public StreamReader CreateMacroStreamReader()
        {
            string path = Path.ChangeExtension(this.name, "js");
            using (StreamWriter sw = new StreamWriter(path))
            {
                sw.WriteLine("dte.ExecuteCommand('File.NewFile');");
            }

            return new StreamReader(path);
        }
    }
}
