//-----------------------------------------------------------------------
// <copyright file="Resources.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace VSMacros.Helpers
{
    public static class Resources
    {
        public static string ValidateErrorInvalidOperation = "You tried to do an invalid operation";
        public static string ValidateErrorStringEmpty = "The string you gave is empty";
        public static string ValidateErrorGuidEmpty = "The Guid is empty";
        public static string ValidateErrorStringWhiteSpace = "The string only contains white space";
        public static string InvalidPIDArgument = "The first command line argument to the execution engine must be an integer representing the PID of the host process, but was instead {0}.";
        public static string InvalidIterationsArgument = "The second command line argument to the execution engine must be an integer representing the number of iterations of execution of the macro, but was instead {0}.";
    }
}
