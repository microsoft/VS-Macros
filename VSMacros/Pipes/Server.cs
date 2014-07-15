//-----------------------------------------------------------------------
// <copyright file="Server.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Diagnostics;
using System.IO.Pipes;
using System.Text;
using System.Threading;
using System.Windows;
using VSMacros.Engines;
using VSMacros.Enums;

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
            Server.ServerStream = new NamedPipeServerStream(guid, PipeDirection.InOut, 1, PipeTransmissionMode.Byte, PipeOptions.Asynchronous);
        }

        #region Getting

        public static string GetMessageFromStream(NamedPipeServerStream serverStream, int sizeOfMessage)
        {
            byte[] messageBuffer = new byte[sizeOfMessage];
            serverStream.Read(messageBuffer, 0, sizeOfMessage);
            return System.Text.UnicodeEncoding.Unicode.GetString(messageBuffer);
        }

        public static int GetIntFromStream(NamedPipeServerStream serverStream)
        {
            byte[] number = new byte[sizeof(int)];
            serverStream.Read(number, 0, sizeof(int));
            return BitConverter.ToInt32(number, 0);
        }

        public static void WaitForMessage()
        {
            bool willKeepRunning = true;

            while (willKeepRunning)
            {
                int typeOfMessage = Server.GetIntFromStream(Server.ServerStream);

                switch ((Packet)typeOfMessage)
                {
                    case Packet.Close:
                        MessageBox.Show("Received a close packet in Server.");
                        Executor.IsEngineInitialized = false;
                        willKeepRunning = false;
                        break;

                    case Packet.Success:
                        Server.HandlePacketSuccess(Server.ServerStream);
                        break;

                    case Packet.ScriptError:
                        Server.HandlePacketScriptError(Server.ServerStream);
                        break;
                }
            }
        }

        private static void HandlePacketScriptError(NamedPipeServerStream serverStream)
        {
            int lineNumber = GetIntFromStream(serverStream);
            int characterPos = GetIntFromStream(serverStream);
            int sizeOfSource = GetIntFromStream(serverStream);
            string source = GetMessageFromStream(serverStream, sizeOfSource);
            int sizeOfDescription = GetIntFromStream(serverStream);
            string description = GetMessageFromStream(serverStream, sizeOfDescription);

            var exceptionMessage = string.Format("{0}: {1} at line {2}, character position {3}.", source, description, lineNumber, characterPos);
            MessageBox.Show(exceptionMessage);
        }

        private static void HandlePacketSuccess(NamedPipeServerStream namedPipeServerStream)
        {
            //throw new NotImplementedException();
        }

        #endregion

        #region Sending

        private static byte[] PackageFilePathMessage(int it, string line)
        {
            byte[] serializedTypeLength = BitConverter.GetBytes((int)Packet.FilePath);
            byte[] serializedIterations = BitConverter.GetBytes((int)it);
            byte[] serializedMessage = UnicodeEncoding.Unicode.GetBytes(line);
            byte[] serializedLength = BitConverter.GetBytes(serializedMessage.Length);

            int type = sizeof(int), iterations = sizeof(int), messageSize = sizeof(int);
            byte[] packet = new byte[type + iterations + messageSize + serializedMessage.Length];

            int offset = 0;
            serializedTypeLength.CopyTo(packet, offset);

            offset += sizeof(int);
            serializedIterations.CopyTo(packet, offset);

            offset += sizeof(int);
            serializedLength.CopyTo(packet, offset);
            
            offset += sizeof(int);
            serializedMessage.CopyTo(packet, offset);

            return packet;
        }

        public static byte[] PackageCloseMessage()
        {
            byte[] serializedType = BitConverter.GetBytes((int)Packet.Close);
            byte[] packet = new byte[sizeof(int)];

            serializedType.CopyTo(packet, 0);

            return packet;
        }

        public static void SendMessageToClient(NamedPipeServerStream serverStream, byte[] packet)
        {
            serverStream.Write(packet, 0, packet.Length);
        }

        internal static void SendCloseRequest()
        {
            byte[] closePacket = PackageCloseMessage();
            SendMessageToClient(Server.ServerStream, closePacket);
        }

        internal static void SendFilePath(int iterations, string path)
        {
            byte[] filePathPacket = PackageFilePathMessage(iterations, path);
            SendMessageToClient(Server.ServerStream, filePathPacket);
        }
    }

        #endregion
}
