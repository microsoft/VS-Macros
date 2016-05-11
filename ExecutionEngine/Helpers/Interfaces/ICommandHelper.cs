//-----------------------------------------------------------------------
// <copyright file="ICommandHelper.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Runtime.InteropServices;

namespace ExecutionEngine.Helpers
{
    [Guid("90d82b99-f713-4663-a0f2-cc72210e57ef")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [ComVisible(true)]
    public interface ICommandHelper
    {
        void DispatchCommandByName(object canonicalName);
        void DispatchCommand(object commandSet, object commandId);
        void DispatchCommandWithArgs(object commandSet, object commandId, ref object pvaIn);
        void ShowMessage(object message);
    }
}
