//-----------------------------------------------------------------------
// <copyright file="RecordedActionBase.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.IO;

namespace VSMacros
{
    internal abstract class RecordedActionBase
    {
        internal abstract void ConvertToJavascript(StreamWriter outputStream);
    }
}