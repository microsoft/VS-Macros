//-----------------------------------------------------------------------
// <copyright file="Program.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using ExecutionEngine.Enums;
using ExecutionEngine.Helpers;
using Microsoft.Internal.VisualStudio.Shell;
using VSMacros.ExecutionEngine.Pipes;

namespace ExecutionEngine
{
    internal class Program
    {
        private static Engine engine;
        private static ParsedScript parsedScript;
        internal const string MacroName = "currentScript";

        internal static void RunMacro(string script, int iterations)
        {
            Validate.IsNotNullAndNotEmpty(script, "script");
            Program.parsedScript = Program.engine.Parse(script);

            for (int i = 0; i < iterations; i++)
            {
                if (!Program.parsedScript.CallMethod(Program.MacroName))
                {
                    if (Site.RuntimeError)
                    {
                        uint activeDocumentModification = 1;
                        var e = Site.RuntimeException;
                        uint modifiedLineNumber = e.Line - activeDocumentModification;

                        byte[] scriptErrorMessage = Client.PackageScriptError(modifiedLineNumber, e.CharacterPosition, e.Source, e.Description);
                        string message = Encoding.Unicode.GetString(scriptErrorMessage);
                        Client.SendMessageToServer(Client.ClientStream, scriptErrorMessage);
                    }
                    else
                    {
                        var e = Site.InternalVSException;
                        byte[] criticalErrorMessage = Client.PackageCriticalError(e.Message, e.Source, e.StackTrace, e.TargetSite.ToString());
                        Client.SendMessageToServer(Client.ClientStream, criticalErrorMessage);
                    }
                    Site.ResetError();
                    break;
                }
            }
        }

        private static void HandleInput()
        {
            int typeOfMessage = Client.GetInt(Client.ClientStream);

            // I know a switch statement seems useless but just preparing for the possibility of other packets.
            switch ((Packet)typeOfMessage)
            {
                case Packet.FilePath:
                    HandleFilePath();
                    break;
            }
        }

        private static void HandleFilePath()
        {
            int iterations = Client.GetInt(Client.ClientStream);
            string message = Client.GetFilePath(Client.ClientStream);
            string unwrappedScript = InputParser.ExtractScript(message);
            string wrappedScript = InputParser.WrapScript(unwrappedScript);
            Program.RunMacro(wrappedScript, iterations);
        }

        internal static Thread CreateReadingThread(int pid)
        {
            Thread readThread = new Thread(() =>
            {
                try
                {
                    Program.engine = new Engine(pid);
                    while (true)
                    {
                        HandleInput();
                    }
                }
                catch (Exception e)
                {
                    if (Client.ClientStream.IsConnected)
                    {
                        byte[] criticalError = Client.PackageCriticalError(e.Message, e.Source, e.StackTrace, e.TargetSite.ToString());
                        Client.SendMessageToServer(Client.ClientStream, criticalError);
                    }
                    else
                    {
#if DEBUG
                        MessageBox.Show("Execution engine's pipe is not connected.");
#endif
                    }
                }
                finally
                {
                    if (Client.ClientStream.IsConnected)
                    {
                        Client.ShutDownServer(Client.ClientStream);
                        Client.ClientStream.Close();
                    }
                    else
                    {
#if DEBUG
                        MessageBox.Show("Execution engine's pipe is not connected.");
#endif
                    }
                }
            });

            readThread.SetApartmentState(ApartmentState.STA);
            return readThread;
        }

        private static void RunFromPipe(string[] separatedArgs)
        {
            Guid guid = InputParser.GetGuid(separatedArgs[0]);
            Client.InitializePipeClientStream(guid);

            int pid = InputParser.GetPid(separatedArgs[1]);
            Thread readThread = CreateReadingThread(pid);
            readThread.Start();
        }

        internal static void Main(string[] args)
        {
            try
            {
                string[] separatedArgs = InputParser.SeparateArgs(args);
                RunFromPipe(separatedArgs);
            }
            catch (Exception e)
            {
                byte[] criticalError = Client.PackageCriticalError(e.Message, e.Source, e.StackTrace, e.TargetSite.ToString());
                Client.SendMessageToServer(Client.ClientStream, criticalError);
            }
        }
    }
}
