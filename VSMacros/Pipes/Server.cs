//-----------------------------------------------------------------------
// <copyright file="Server.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading;
using System.Windows;
using VisualStudio.Macros.ExecutionEngine.Pipes;
using VSMacros.Engines;
using VSMacros.Enums;
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

        #region Getting

        public static void WaitForMessage()
        {
            Server.ServerStream.WaitForConnection();

            var formatter = new BinaryFormatter();
            formatter.Binder = new BinderHelper();

            bool shouldKeepRunning = true;
            while (shouldKeepRunning && Server.ServerStream.IsConnected)
            {
                try
                {
                    var type = (PacketType)formatter.Deserialize(Server.ServerStream);

                    switch (type)
                    {
                        case PacketType.Empty:
#if DEBUG
                        Manager.Instance.ShowMessageBox("Pipes are no longer connected.");
#endif
                            Executor.IsEngineInitialized = false;
                            shouldKeepRunning = false;
                            break;

                        case PacketType.Close:
#if DEBUG
                        Manager.Instance.ShowMessageBox("Received a close packet in Server.");
#endif
                            Executor.IsEngineInitialized = false;
                            shouldKeepRunning = false;
                            break;

                        case PacketType.Success:
                            var executor = Manager.Instance.executor;
                            executor.SendCompletionMessage(isError: false, errorMessage: string.Empty);
                            break;

                        case PacketType.ScriptError:
                            executor = Manager.Instance.executor;
                            string error = Server.GetScriptError(Server.ServerStream);
                            executor.SendCompletionMessage(isError: true, errorMessage: error);
                            break;

                        case PacketType.CriticalError:
                            error = Server.GetCriticalError(Server.ServerStream);
#if DEBUG
                        Manager.Instance.ShowMessageBox(error);
#endif
                            Server.ServerStream.Close();

                            if (Executor.Job != null)
                            {
                                Executor.Job.Close();
                            }

                            Executor.IsEngineInitialized = false;
                            shouldKeepRunning = false;
                            break;
                    }
                }
                catch (System.Runtime.Serialization.SerializationException e)
                {
#if DEBUG
                    Debug.WriteLine("Server has shut down: " + e.Message);
                    // TODO: What else do I need to do here?
#endif
                }
            } 
        }

        private static string GetScriptError(NamedPipeServerStream serverStream)
        {
            var formatter = new BinaryFormatter();
            formatter.Binder = new BinderHelper();
            var scriptError = (ScriptError)formatter.Deserialize(Server.ServerStream);

            int lineNumber = scriptError.LineNumber;
            int column = scriptError.Column;
            string source = scriptError.Source ?? "Script Error";
            string description = scriptError.Description ?? "Command not valid in this context";
            string period = description[description.Length - 1] == '.' ? string.Empty : ".";

            var exceptionMessage = string.Format("{0}{3}{3}Line Number: {1}{3}Cause: {2}{4}", source, lineNumber, description, Environment.NewLine, period);
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

        #endregion

        #region Sending

        public static void SendMessageToClient(NamedPipeServerStream serverStream, byte[] packet)
        {
            try
            {
                serverStream.Write(packet, 0, packet.Length);
            }
            catch (OperationCanceledException e)
            {
                // TODO: THis needs to be preserved elsewhere.
#if DEBUG
                Manager.Instance.ShowMessageBox(string.Format("The server thread was terminated.\n\n{0}: {1}\n{2}{3}", e.Source, e.Message, e.TargetSite.ToString(), e.StackTrace));
#endif
                VSMacrosPackage.Current.ClearStatusBar();
            }
            
        }

        internal static void SendFilePath(int iterations, string path)
        {
            var type = PacketType.FilePath;
            BinaryFormatter formatter = new BinaryFormatter();
            formatter.Serialize(Server.ServerStream, type);
            Debug.WriteLine("sent type at: " + DateTime.Now);

            var filePath = new FilePath();
            filePath.Iterations = iterations;
            filePath.Path = path;
            formatter.Serialize(Server.ServerStream, filePath);
            Debug.WriteLine("sent file path at: " + DateTime.Now);
        }
    }
        #endregion
}
