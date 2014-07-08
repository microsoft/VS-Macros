using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using VSMacros.ExecutionEngine;

namespace VSMacros.ExecutionEngine.Helpers
{
    public static class InputParser
    {
        internal static bool IsRequestToClose(string s)
        {
            return s[0] == '@';
        }

        internal static bool IsDebuggerStopped(string message)
        {
            return string.IsNullOrEmpty(message);
        }

        internal static string[] SeparateArgs(string[] args)
        {
            string[] stringSeparator = new string[] { "[delimiter]" };
            string[] separatedArgs = args[0].Split(stringSeparator, StringSplitOptions.RemoveEmptyEntries);
            return separatedArgs;
        }

        internal static int GetPid(string unparsedPid)
        {
            int pid;

            Validate.IsNotNullAndNotEmpty(unparsedPid, "unparsedPid");

            if (!int.TryParse(unparsedPid, out pid))
            {
                MessageBox.Show("The pid is invalid.");
                throw new ArgumentException(string.Format(CultureInfo.CurrentUICulture, Resources.InvalidPIDArgument, unparsedPid, "unparsedPid"));
            }

            return pid;
        }

        internal static string GetGuid(string guid)
        {
            Validate.IsNotNullAndNotEmpty(guid, "guid");
            return guid;
        }

        internal static string DecodePath(string encodedPath)
        {
            return encodedPath.Replace("%20", " ");
        }

        internal static short GetNumberOfIterations(string iter)
        {
            short iterations;

            Validate.IsNotNullAndNotEmpty(iter, "iter");

            if (!short.TryParse(iter, out iterations))
            {
                throw new ArgumentException(string.Format(CultureInfo.CurrentUICulture, Resources.InvalidIterationsArgument, iter, "iter"));
            }

            return iterations;
        }

        internal static string ExtractScript(string path)
        {
            Validate.IsNotNullAndNotEmpty(path, "path");
            return File.ReadAllText(path);
        }

        internal static string WrapScript(string unwrapped)
        {
            string wrapped = "function currentScript() {";
            wrapped += unwrapped;
            wrapped += "}";

            return wrapped;
        }

        internal static string RemoveComments(string commented)
        {
            var re = @"(@(?:""[^""]*"")+|""(?:[^""\n\\]+|\\.)*""|'(?:[^'\n\\]+|\\.)*')|//.*|/\*(?s:.*?)\*/";
            var noComments = Regex.Replace(commented, re, "$1");
            Console.WriteLine(noComments);
            return noComments;
        }
    }
}
