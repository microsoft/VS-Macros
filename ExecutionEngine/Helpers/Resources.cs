using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VSMacros.ExecutionEngine
{
    public static class Resources
    {
        public static string ValidateError_InvalidOperation = "Stub for now: You tried to do an invalid operation";
        public static string ValidateError_StringEmpty = "Stub for now: The string you gave is empty";
        public static string ValidateError_GuidEmpty = "Stub for now: The Guid is empty";
        public static string ValidateError_StringWhiteSpace = "Stub for now: The string only contains white space";
        public static string InvalidPIDArgument = "The first command line argument to the execution engine must be an integer representing the PID of the host process, but was instead {0}.";
        public static string InvalidIterationsArgument = "The second command line argument to the execution engine must be an integer representing the number of iterations of execution of the macro, but was instead {0}.";
    }
}
