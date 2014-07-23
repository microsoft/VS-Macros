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
//using VisualStudio.Macros.ExecutionEngine.Pipes;
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

        //public static string GetMessageFromStream(NamedPipeServerStream serverStream, int sizeOfMessage)
        //{
        //    byte[] messageBuffer = new byte[sizeOfMessage];
        //    serverStream.Read(messageBuffer, 0, sizeOfMessage);
        //    return Encoding.Unicode.GetString(messageBuffer);
        //}

        //public static int GetIntFromStream(NamedPipeServerStream serverStream)
        //{
        //    if (serverStream.IsConnected)
        //    {
        //        byte[] number = new byte[sizeof(int)];
        //        serverStream.Read(number, 0, sizeof(int));
        //        return BitConverter.ToInt32(number, 0);
        //    }
        //    else
        //    {
        //        return -1;
        //    }
        //}

        public static void WaitForMessage()
        {
            Server.ServerStream.WaitForConnection();

            var formatter = new BinaryFormatter();
            var binder = new BinderHelper();
            formatter.Binder = binder;

            bool shouldKeepRunning = true;
            while (shouldKeepRunning)
            {
                try
                {
                    var type = (PacketType)formatter.Deserialize(Server.ServerStream);

                    //int typeOfMessage = Server.GetIntFromStream(Server.ServerStream);

                    switch (type)
                    //switch ((Packet)typeOfMessage)
                    {
                        //case Packet.Empty:
                        case PacketType.Empty:
#if DEBUG
                        Manager.Instance.ShowMessageBox("Pipes are no longer connected.");
#endif
                            Executor.IsEngineInitialized = false;
                            shouldKeepRunning = false;
                            break;

                        //case Packet.Close:
                        case PacketType.Close:
#if DEBUG
                        Manager.Instance.ShowMessageBox("Received a close packet in Server.");
#endif
                            Executor.IsEngineInitialized = false;
                            shouldKeepRunning = false;
                            break;

                        //case Packet.Success:
                        case PacketType.Success:
                            var executor = Manager.Instance.executor;
                            executor.SendCompletionMessage(isError: false, errorMessage: string.Empty);
                            break;

                        //case Packet.ScriptError:
                        case PacketType.ScriptError:
                            executor = Manager.Instance.executor;
                            string error = Server.GetScriptError(Server.ServerStream);
                            executor.SendCompletionMessage(isError: true, errorMessage: error);
                            break;

                        //case Packet.CriticalError:
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
                    MessageBox.Show(e.Message);
#endif
                    // TODO: What else do I need to do here?
                    // It lags when VS is shut down
                }
            } 
        }

        private static string GetScriptError(NamedPipeServerStream serverStream)
        {
            var formatter = new BinaryFormatter();
            var binder = new BinderHelper();
            formatter.Binder = binder;
            var scriptError = (ScriptError)formatter.Deserialize(Server.ServerStream);

            //int lineNumber = GetIntFromStream(serverStream);
            int lineNumber = scriptError.LineNumber;
            //int column = GetIntFromStream(serverStream);
            int column = scriptError.Column;
            //int sizeOfSource = GetIntFromStream(serverStream);
            //string source = GetMessageFromStream(serverStream, sizeOfSource);
            string source = scriptError.Source;
            //int sizeOfDescription = GetIntFromStream(serverStream);
            //string description = GetMessageFromStream(serverStream, sizeOfDescription);
            string description = scriptError.Description;

            var exceptionMessage = string.Format("{0}: {1} at line {2}, column {3}.", source, description, lineNumber, column);
            return exceptionMessage;
        }

        private static string GetCriticalError(NamedPipeServerStream serverStream)
        {
            var formatter = new BinaryFormatter();
            var criticalError = (CriticalError)formatter.Deserialize(Server.ServerStream);

            //int messageSize = GetIntFromStream(serverStream);
            //string message = GetMessageFromStream(serverStream, messageSize);
            string message = criticalError.Message;
            //int sourceSize = GetIntFromStream(serverStream);
            //string source = GetMessageFromStream(serverStream, sourceSize);
            string source = criticalError.StackTrace;
            //int stackTraceSize = GetIntFromStream(serverStream);
            //string stackTrace = GetMessageFromStream(serverStream, stackTraceSize);
            string stackTrace = criticalError.StackTrace;
            //int targetSiteSize = GetIntFromStream(serverStream);
            //string targetSite = GetMessageFromStream(serverStream, targetSiteSize);
            string targetSite = criticalError.TargetSite;

            var exceptionMessage = string.Format("{0}: {1}{2}Stack Trace: {3}{2}{2}TargetSite:{4}", source, message, Environment.NewLine, stackTrace, targetSite);
            return exceptionMessage;
        }

        #endregion

        #region Sending

        //private static byte[] PackageFilePathMessage(int it, string line)
        //{

        //    byte[] serializedTypeLength = BitConverter.GetBytes((int)Packet.FilePath);
        //    byte[] serializedIterations = BitConverter.GetBytes((int)it);
        //    byte[] serializedMessage = Encoding.Unicode.GetBytes(line);
        //    byte[] serializedLength = BitConverter.GetBytes(serializedMessage.Length);

        //    int type = sizeof(int), iterations = sizeof(int), messageSize = sizeof(int);
        //    byte[] packet = new byte[type + iterations + messageSize + serializedMessage.Length];

        //    int offset = 0;
        //    serializedTypeLength.CopyTo(packet, offset);

        //    offset += type;
        //    serializedIterations.CopyTo(packet, offset);

        //    offset += iterations;
        //    serializedLength.CopyTo(packet, offset);
            
        //    offset += messageSize;
        //    serializedMessage.CopyTo(packet, offset);

        //    return packet;
        //}

        //public static byte[] PackageCloseMessage()
        //{
        //    return BitConverter.GetBytes((int)Packet.Close);
        //}

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

            var filePath = new FilePath();
            filePath.Iterations = iterations;
            filePath.Path = path;
            formatter.Serialize(Server.ServerStream, filePath);

            //byte[] filePathPacket = PackageFilePathMessage(iterations, path);
            //SendMessageToClient(Server.ServerStream, filePathPacket);
        }
    }
        #endregion
}
