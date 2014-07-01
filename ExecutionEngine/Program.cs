//-----------------------------------------------------------------------
// <copyright file="Program.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ExecutionEngine
{
    internal class Program
    {
        private static Engine engine;
        private static ParsedScript parsedScript;
        //private static int pid;
        private static string macroName = "currentScript";

        // Helper methods
        private static string[] SeparateArgs(string[] args)
        {
            string[] stringSeparator = new string[] {"[delimiter]"};
            string[] separatedArgs = args[0].Split(stringSeparator, StringSplitOptions.RemoveEmptyEntries);
            return separatedArgs;
        }

        private static short DetermineNumberOfIterations(string iter)
        {
            short iterations;

            if (string.IsNullOrEmpty(iter) || !short.TryParse(iter, out iterations))
            {
                throw new ArgumentException(iter);
            }

            return iterations;
        }

        private static string ExtractScript(string path)
        {
            if (!string.IsNullOrEmpty(path))
            {
                return File.ReadAllText(path);
            }
            else
            {
                throw new ArgumentException(path);
            }
        }

        private static string WrapScript(string unwrapped)
        {
            string wrapped = "function currentScript() {";
            wrapped += unwrapped;
            wrapped += "}";

            return wrapped;
        }

        internal static void RunMacro(string script, int iterations = 1)
        {
            if (!string.IsNullOrEmpty(script))
            {
                Program.parsedScript = Program.engine.Parse(script);
            }
            else
            {
                throw new ArgumentException(script);
            }

            for (int i = 0; i < iterations; i++)
            {
                Program.parsedScript.CallMethod(Program.macroName);
            }
        }

        internal static void RunFromExtension(string[] separatedArgs)
        {
            string unparsedPid = separatedArgs[0];
            string unparsedIter = separatedArgs[1];
            string encodedPath = separatedArgs[2];

            int pid = DeterminePid(unparsedPid);
            short iterations = DetermineNumberOfIterations(unparsedIter);
            string decodedPath = DecodePath(encodedPath);
            string unwrappedScript = ExtractScript(decodedPath);
            string wrappedScript = WrapScript(unwrappedScript);

            Program.engine = new Engine(pid);
            RunMacro(wrappedScript, iterations);
        }

        private static int DeterminePid(string unparsedPid)
        {
            int pid;
            if (!int.TryParse(unparsedPid, out pid))
            {
                throw new ArgumentException(unparsedPid);
            }
            return pid;
        }

        private static string DecodePath(string encodedPath)
        {
            return encodedPath.Replace("%20", " ");
        }

        internal static void RunAsStartupProject(int tempPid)
        {
            Program.engine = new Engine(tempPid);

            string unwrapped = File.ReadAllText(@"C:\Users\t-grawa\Desktop\test.js");
            string wrapped = WrapScript(unwrapped);

            RunMacro(wrapped, 1);
        }

        internal static void Main(string[] args)
        {
            Console.WriteLine("Hello there!  Welcome to our macro extension!");

            if (args.Length > 0)
            {
                //MessageBox.Show("Attach debugger here");
                string[] separatedArgs = SeparateArgs(args);
                RunFromExtension(separatedArgs);
            }
            else
            {
                short pidOfCurrentDevenv = 10588;
                RunAsStartupProject(pidOfCurrentDevenv);   
            }
        }
    }
}
