//-----------------------------------------------------------------------
// <copyright file="ScriptState.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------


namespace ExecutionEngine.Enums
{
    internal enum ScriptState
    {
        Uninitialized = 0,
        Started = 1,
        Connected = 2,
        Disconnected = 3,
        Closed = 4,
        Initialized = 5
    }
}
