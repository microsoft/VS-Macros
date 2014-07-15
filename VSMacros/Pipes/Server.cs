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

                // TODO: I think FilePath needs to be modified as well

                switch ((Packet)typeOfMessage)
                {
                    case Packet.FilePath:
                        string message = ParseFilePath(Server.ServerStream);
                        break;

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

                    case (Packet.Iterations):
                        Server.HandlePacketOtherError(Server.ServerStream);
                        break;
                }
            }
        }

        private static void HandlePacketOtherError(NamedPipeServerStream namedPipeServerStream)
        {
            throw new NotImplementedException();
        }

        private static void HandlePacketScriptError(NamedPipeServerStream serverStream)
        {
            int lineNumber = GetIntFromStream(serverStream);
            int sizeOfSource = GetIntFromStream(serverStream);
            string source = GetMessageFromStream(serverStream, sizeOfSource);
            int sizeOfDescription = GetIntFromStream(serverStream);
            string description = GetMessageFromStream(serverStream, sizeOfDescription);

            var exceptionMessage = string.Format("{0}: {1} at line {2}.", source, description, lineNumber);
            string gloat = "Visual Studio presents:\n";
            MessageBox.Show(gloat + exceptionMessage);
        }

        private static void HandlePacketSuccess(NamedPipeServerStream namedPipeServerStream)
        {
            //throw new NotImplementedException();
        }

        private static string ParseFilePath(NamedPipeServerStream serverStream)
        {
            int sizeOfMessage = Server.GetIntFromStream(serverStream);
            string message = Server.GetMessageFromStream(serverStream, sizeOfMessage);
            return message;
        }

        #endregion

        private static byte[] PackageFilePathMessage(string line)
        {
            byte[] serializedTypeLength = BitConverter.GetBytes((int)Packet.FilePath);

            byte[] serializedMessage = UnicodeEncoding.Unicode.GetBytes(line);
            int message = serializedMessage.Length;

            byte[] serializedLength = BitConverter.GetBytes(message);

            int type = sizeof(int), messageSize = sizeof(int);
            byte[] packet = new byte[type + messageSize + message];

            serializedTypeLength.CopyTo(packet, 0);

            int offset = sizeof(int);
            serializedLength.CopyTo(packet, offset);
            offset += sizeof(int);

            serializedMessage.CopyTo(packet, offset);

            return packet;
        }

        private static byte[] PackageFilePathMessage(int iterations, string line)
        {
            byte[] serializedTypeLength = BitConverter.GetBytes((int)Packet.FilePath);
            byte[] serializedIterations = BitConverter.GetBytes((int)iterations);
            byte[] serializedMessage = UnicodeEncoding.Unicode.GetBytes(line);
            byte[] serializedLength = BitConverter.GetBytes(serializedMessage.Length);

            int typePlaceholder = sizeof(int), iterationsPlaceholder = sizeof(int), messageSizePlaceholder = sizeof(int);
            byte[] packet = new byte[typePlaceholder + iterationsPlaceholder + messageSizePlaceholder + serializedMessage.Length];

            int offset = 0;
            serializedTypeLength.CopyTo(packet, offset);

            offset += sizeof(int);
            serializedIterations.CopyTo(packet, offset);

            offset += sizeof(int);
            serializedLength.CopyTo(packet, offset);
            
            offset += sizeof(int);
            serializedMessage.CopyTo(packet, offset);

            Debug.WriteLine("This is what the server is sending oout:");
            for (int i = 0; i < packet.Length; i++)
            {
                Debug.Write(packet[i]);
                Debug.Write(' ');
            }

                return packet;
        }

        public static byte[] PackageCloseMessage()
        {
            byte[] serializedType = BitConverter.GetBytes((int)Packet.Close);

            int type = sizeof(int);
            byte[] packet = new byte[type];

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

        internal static void SendFilePath(string path)
        {
            byte[] filePathPacket = PackageFilePathMessage(path);
            SendMessageToClient(Server.ServerStream, filePathPacket);
        }

        internal static void SendFilePath(int iterations, string path)
        {
            byte[] filePathPacket = PackageFilePathMessage(iterations, path);
            SendMessageToClient(Server.ServerStream, filePathPacket);
        }
    }
}
