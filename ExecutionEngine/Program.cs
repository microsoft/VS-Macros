//-----------------------------------------------------------------------
// <copyright file="Program.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Runtime.Serialization.Formatters.Binary;
using System.Threading;
using System.Windows.Forms;
using ExecutionEngine.Helpers;
using Microsoft.Internal.VisualStudio.Shell;
using VisualStudio.Macros.ExecutionEngine.Pipes;
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

            int i = 0;
            for (i = 0; i < iterations; i++)
            {
                if (!Program.parsedScript.CallMethod(Program.MacroName))
                {
                    if (Site.RuntimeError)
                    {
                        uint activeDocumentModification = 1;
                        var e = Site.RuntimeException;
                        uint modifiedLineNumber = e.Line - activeDocumentModification;

                        Client.SendScriptError(modifiedLineNumber, e.CharacterPosition, e.Source, e.Description);
                    }
                    else
                    {
                        var e = Site.InternalVSException;
                        Client.SendCriticalError(e.Message, e.Source, e.StackTrace, e.TargetSite.ToString());
                    }
                    Site.ResetError();
                    break;
                }
            }

            Client.SendSuccessMessage();
        }

        private static void HandleInput()
        {
            BinaryFormatter formatter = new BinaryFormatter();
            formatter.Binder = new BinderHelper();
            var type = (PacketType)formatter.Deserialize(Client.ClientStream);

            // I know a switch statement seems useless but just preparing for the possibility of other packets.
            switch (type)
            {
                case PacketType.FilePath:
                    HandleFilePath();
                    break;
            }
        }

        private static void HandleFilePath()
        {
            // Just make one static formatter
            var formatter = new BinaryFormatter();
            var filePath = (FilePath)formatter.Deserialize(Client.ClientStream);

            int iterations = filePath.Iterations;
            string message = filePath.Path;
            string unwrappedScript = InputParser.ExtractScript(message);
            string wrappedScript = InputParser.WrapScript(unwrappedScript);
            Program.RunMacro(wrappedScript, iterations);
        }

        internal static Thread CreateReadingThread(int pid, string version)
        {
            Thread readThread = new Thread(() =>
            {
                try
                {
                    Program.engine = new Engine(pid, version);
                    while (true)
                    {
                        HandleInput();
                    }
                }
                catch (Exception e)
                {
                    if (Client.ClientStream.IsConnected)
                    {
                        Client.SendCriticalError(e.Message, e.Source, e.StackTrace, e.TargetSite.ToString());
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
            string version = separatedArgs[2];

            Thread readThread = CreateReadingThread(pid, version);
            readThread.Start();
        }

        internal static void Main(string[] args)
        {
            try
            {
                //MessageBox.Show("hello");
                string[] separatedArgs = InputParser.SeparateArgs(args);
                RunFromPipe(separatedArgs);
            }
            catch (Exception e)
            {
                Client.SendCriticalError(e.Message, e.Source, e.StackTrace, e.TargetSite.ToString());
            }
        }
    }
}
