using System;

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
