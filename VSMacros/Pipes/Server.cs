using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO.Pipes;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using VSMacros.Pipes;
using VSMacros.Engines;
using VSMacros.Enums;
using System.Windows;

namespace VSMacros.Pipes
{
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
            //string pipeGuid = Server.Guid.ToString();

            bool willKeepRunning = true;

            while (willKeepRunning)
            {
                int typeOfMessage = Server.GetIntFromStream(Server.ServerStream);

                switch ((Packet)typeOfMessage)
                {
                    case (Packet.FilePath):
                        string message = ParseFilePath(Server.ServerStream);
                        Debug.WriteLine("message: " + message);
                        break;

                    case (Packet.Close):
                        MessageBox.Show("in Server, Packet.Close");
                        Executor.IsEngineInitialized = false;
                        willKeepRunning = false;
                        break;

                    case (Packet.Success):
                        Server.HandlePacketSuccess(Server.ServerStream);
                        break;

                    case (Packet.ScriptError):
                        Server.HandlePacketScriptError(Server.ServerStream);
                        break;

                    case (Packet.OtherError):
                        Server.HandlePacketOtherError(Server.ServerStream);
                        break;
                }
            }
        }

        private static void HandlePacketOtherError(NamedPipeServerStream namedPipeServerStream)
        {
            throw new NotImplementedException();
        }

        private static void HandlePacketScriptError(NamedPipeServerStream namedPipeServerStream)
        {
            throw new NotImplementedException();
        }

        private static void HandlePacketSuccess(NamedPipeServerStream namedPipeServerStream)
        {
            throw new NotImplementedException();
        }

        private static string ParseFilePath(NamedPipeServerStream serverStream)
        {
            int sizeOfMessage = Server.GetIntFromStream(serverStream);
            string message = Server.GetMessageFromStream(serverStream, sizeOfMessage);
            return message;
        }

        #endregion

        #region Sending

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

        internal static void SendFilePath(string path)
        {
            byte[] filePathPacket = PackageFilePathMessage(path);
            SendMessageToClient(Server.ServerStream, filePathPacket);
        }

        #endregion
    }
}
