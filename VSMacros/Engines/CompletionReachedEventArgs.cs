using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

// TODO: I am not sure if this is the right place to define this class.
// For now I will just keep it closer to the class that uses it.

namespace VSMacros.Engines
{
    class CompletionReachedEventArgs : EventArgs
    {
        public bool Success { get; set; }
    }
}
