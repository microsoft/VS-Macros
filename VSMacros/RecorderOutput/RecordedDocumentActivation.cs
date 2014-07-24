//-----------------------------------------------------------------------
// <copyright file="RecordedDocumentActivation.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.IO;

namespace VSMacros.RecorderOutput
{
    internal sealed class RecordedDocumentActivation : RecordedActionBase
    {
        private string docPath;

        public RecordedDocumentActivation(string path)
        {
            this.docPath = path;
        }
        internal override void ConvertToJavascript(StreamWriter outputStream)
        {
            string escapedInput = this.docPath.Replace("\\", "\\\\");
            outputStream.WriteLine("dte.Documents.Item(\"" + escapedInput + "\").Activate();");
        }
    }
}
