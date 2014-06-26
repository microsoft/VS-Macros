using System;

namespace ExecutionEngine.Enums
{
    [Flags]
    internal enum ScriptText
    {
        None = 0,
        DelayExecution = 1,
        IsVisible = 2,
        IsExpression = 32,
        IsPersistent = 64,
        HostManageSource = 128
    }
}
