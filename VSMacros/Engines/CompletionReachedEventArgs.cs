//-----------------------------------------------------------------------
// <copyright file="CompletionReachedEventArgs.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;

// TODO: I am not sure if this is the right place to define this class.
// For now I will just keep it closer to the class that uses it.

namespace VSMacros.Engines
{
    public class CompletionReachedEventArgs : EventArgs
    {
        public CompletionReachedEventArgs(bool success, string errorMessage)
        {
            this.IsError = success;
            this.ErrorMessage = errorMessage;
        }

        public bool IsError { get; set; }
        public string ErrorMessage { get; set; }
    }
}
