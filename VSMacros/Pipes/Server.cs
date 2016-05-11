//-----------------------------------------------------------------------
// <copyright file="Server.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.IO.Pipes;
using System.Runtime.Serialization.Formatters.Binary;
using System.Threading;
using System.Windows.Threading;
using VisualStudio.Macros.ExecutionEngine.Pipes;
using VSMacros.Engines;
using VSMacros.ExecutionEngine.Pipes;

namespace VSMacros.Pipes
{
    /// <summary>
    /// Aids in IPC.
    /// </summary>
    public static class Server
    {
        public static NamedPipeServerStream ServerStream;
        public static Guid Guid;
        public static Thread serverWait;
        private static Executor executor = Manager.Instance.executor;
        static Dispatcher dispatcher = Dispatcher.CurrentDispatcher;

        public static void InitializeServer()
        {
            Server.Guid = Guid.NewGuid();
            var guid = Server.Guid.ToString();
            Server.ServerStream = new NamedPipeServerStream(
                pipeName:guid, 
                direction: PipeDirection.InOut, 
                maxNumberOfServerInstances: 1, 
                transmissionMode: PipeTransmissionMode.Byte, 
                options: PipeOptions.Asynchronous);
        }

        private static void SendEngineFailedSilentlyCompletionMessage() {
            Server.executor.SendCompletionMessage(isError: false, errorMessage: string.Empty);
        }

        public static void WaitForMessage()
        {
            Server.ServerStream.WaitForConnection();

            var formatter = new BinaryFormatter();
            formatter.Binder = new BinderHelper();

            bool shouldKeepRunning = true;
            while (shouldKeepRunning)
            {
                if (Executor.executionEngine.HasExited)
                {
                    SendEngineFailedSilentlyCompletionMessage();
                    shouldKeepRunning = false;
                }

                try
                {
                    var type = (PacketType)formatter.Deserialize(Server.ServerStream);

                    switch (type)
                    {
                        case PacketType.Empty:
                            Server.executor.IsEngineRunning = false;
                            shouldKeepRunning = false;
                            break;

                        case PacketType.Close:
                            shouldKeepRunning = false;
                            break;

                        case PacketType.Success:
                            Server.executor.SendCompletionMessage(isError: false, errorMessage: string.Empty);
                            break;

                        case PacketType.GenericScriptError:
                            string error = Server.GetGenericScriptError(Server.ServerStream);
                            Server.executor.SendCompletionMessage(isError: true, errorMessage: error);
                            break;

                        case PacketType.CriticalError:
                            error = Server.GetCriticalError(Server.ServerStream);
                            Server.ServerStream.Close();
                            CloseExecutorJob();
                            shouldKeepRunning = false;
                            break;
                    }
                }
                catch (System.Runtime.Serialization.SerializationException)
                {
                }
            } 
        }

        private static void CloseExecutorJob()
        {
            if (Executor.Job != null)
            {
                Executor.Job.Close();
            }
        }

        private static string GetGenericScriptError(NamedPipeServerStream serverStream)
        {
            var formatter = new BinaryFormatter();
            formatter.Binder = new BinderHelper();
            var scriptError = (GenericScriptError)formatter.Deserialize(Server.ServerStream);

            int lineNumber = scriptError.LineNumber;
            int column = scriptError.Column;
            string source = string.Format(Resources.MacroFileError, Server.executor.CurrentlyExecutingMacro);
            string description = scriptError.Description;
            string period = description[description.Length - 1] == '.' ? string.Empty : ".";

            var exceptionMessage = string.Format("{0}{3}{3}Line {1}: {2}{4}", source, lineNumber, description, Environment.NewLine, period);
            return exceptionMessage;
        }

        private static string GetCriticalError(NamedPipeServerStream serverStream)
        {
            var formatter = new BinaryFormatter();
            var criticalError = (CriticalError)formatter.Deserialize(Server.ServerStream);

            string message = criticalError.Message;
            string source = criticalError.StackTrace;
            string stackTrace = criticalError.StackTrace;
            string targetSite = criticalError.TargetSite;

            var exceptionMessage = string.Format("{0}: {1}{2}Stack Trace: {3}{2}{2}TargetSite:{4}", source, message, Environment.NewLine, stackTrace, targetSite);
            return exceptionMessage;
        }

        internal static void SendFilePath(int iterations, string path)
        {
            try
            {
                var type = PacketType.FilePath;
                BinaryFormatter formatter = new BinaryFormatter();
                formatter.Serialize(Server.ServerStream, type);

                var filePath = new FilePath();
                filePath.Iterations = iterations;
                filePath.Path = path;
                formatter.Serialize(Server.ServerStream, filePath);
            }
            catch (Exception)
            {
            }
        }
    }
}
