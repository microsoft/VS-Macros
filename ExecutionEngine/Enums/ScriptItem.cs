//-----------------------------------------------------------------------
// <copyright file="ScriptItem.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;

namespace ExecutionEngine.Enums
{
    [Flags]
    internal enum ScriptItem
    {
        None = 0,
        IsVisible = 2,
        IsSource = 4,
        GlobalMembers = 8,
        IsPersistent = 64,
        CodeOnly = 512,
        NoCode = 1024
    }
}
