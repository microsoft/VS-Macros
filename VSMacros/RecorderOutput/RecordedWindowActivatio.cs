//-----------------------------------------------------------------------
// <copyright file="RecordedWindowActivation.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.IO;

namespace VSMacros.RecorderOutput
{
    [DebuggerDisplay("{windowName}")]
    internal sealed class RecordedWindowActivation : RecordedActionBase
    {
        private Guid activatedWindow;
        private string windowName;

        public RecordedWindowActivation(Guid guid, string name)
        {
            this.activatedWindow = guid;
            this.windowName = name;
        }
        internal override void ConvertToJavascript(StreamWriter outputStream)
        {
            outputStream.WriteLine("dte.Windows.Item(\"{" + this.activatedWindow.ToString() + "}\").Activate();");
        }
    }
}
