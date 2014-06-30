using System;
using Microsoft.Internal.VisualStudio.Shell;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using System.IO;
using EnvDTE;
using EnvDTE80;
using System.Diagnostics;

namespace VSMacros
{
    internal abstract class RecordedActionBase
    {
        internal abstract void ConvertToJavascript(StreamWriter outputStream);
    }

    [DebuggerDisplay("{commandName}")]
    internal sealed class RecordedCommand : RecordedActionBase
    {
        private Guid commandSetGuid;
        private uint commandId;
        private string commandName;
        private char input; 

        internal RecordedCommand(Guid commandSetGuid, uint commandId, string commandName, char input)
        {
            this.commandName = commandName;
            this.commandSetGuid = commandSetGuid;
            this.commandId = commandId;
            this.input = input;
        }

        internal override void ConvertToJavascript(StreamWriter outputStream)
        {
            // Converts from the GUID / uint to Javascript and writes the result to the given stream here. You may also need a constructor param of type IServiceProvider or the IVsCmdNameMapping object so you can convert it to nice human readable text names (though there is a way to get DTE to raise commands by GUID / uint pair as well
            if (this.input == 0)
            {
                outputStream.WriteLine("dte.Commands.Raise(\"" + this.commandSetGuid + "\", " + this.commandId + ")");
            }
            else
            {
                outputStream.WriteLine("dte.Commands.Raise(\"" + this.commandSetGuid + "\", " + this.commandId + ", " + this.input + ")");
            }
        }
    }

    [DebuggerDisplay("{windowName}")]
    internal sealed class RecordedWindowActivation : RecordedActionBase
    {
        private Guid activatedWindow;
        private string windowName;

        public RecordedWindowActivation(Guid guid,string name)
        {
            this.activatedWindow = guid;
            this.windowName = name;
        }
        internal override void ConvertToJavascript(StreamWriter outputStream)
        {
            // Write the code that uses DTE to activate the appropriate window here and write the output to the given stream
            outputStream.WriteLine("dte.Documents.Item(\"" + this.activatedWindow.ToString() + "\").ActiveWindow.Activate()");
            
        }
    }

    internal sealed class RecordedDocumentActivation : RecordedActionBase
    {
        private string docPath;

        public RecordedDocumentActivation(string path)
        {
            this.docPath = path;
        }
        internal override void ConvertToJavascript(StreamWriter outputStream)
        {
            outputStream.WriteLine("dte.Documents.Item(\"" + this.docPath + "\").ActiveWindow.Activate()");
        }
    }
}