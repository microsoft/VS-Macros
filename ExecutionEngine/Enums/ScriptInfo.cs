//-----------------------------------------------------------------------
// <copyright file="ScriptInfo.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;

namespace ExecutionEngine.Enums
{
    [Flags]
    internal enum ScriptInfo
    {
        None = 0,
        IUnknown = 1,
        ITypeInfo = 2
    }
}
