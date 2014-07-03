//-----------------------------------------------------------------------
// <copyright file="RecordedActionBase.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

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
}