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
            outputStream.WriteLine("dte.Documents.Item(\"" + Path.GetFileName(this.docPath) + "\").Activate()");
        }
    }
}
