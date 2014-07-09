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

        public static string GetMessageFromStream(NamedPipeServerStream namedPipeServerStream, int sizeOfMessage)
        {
            byte[] messageBuffer = new byte[sizeOfMessage];
            namedPipeServerStream.Read(messageBuffer, 0, sizeOfMessage);
            return System.Text.UnicodeEncoding.Unicode.GetString(messageBuffer);
        }

        public static int GetSizeOfMessageFromStream(NamedPipeServerStream namedPipeServerStream)
        {
            int intSize = sizeof(int);
            byte[] sizeBuffer = new byte[intSize];
            namedPipeServerStream.Read(sizeBuffer, 0, intSize);
            return BitConverter.ToInt32(sizeBuffer, 0);
        }

        public static void WaitForMessage()
        {
            string pipeGuid = Server.Guid.ToString();

            while (true)
            {
                int sizeOfMessage = Server.GetSizeOfMessageFromStream(Server.ServerStream);
                string message = Server.GetMessageFromStream(Server.ServerStream, sizeOfMessage);

                if (message.ToLower().Equals("close"))
                {
                    Executor.IsEngineInitialized = false;
                    break;
                }
            }
        }

        public static void SendMessage()
        {
            string rawMessage = "@";
            byte[] message = PackageFilePathMessage(rawMessage);

            if (Server.ServerStream.IsConnected)
            {   
                for (int i = 0; i < message.Length; i++)
                {
                    Debug.Write(message[i]);
                }
                Debug.WriteLine(" ");
                
                SendMessageToClient(Server.ServerStream, message);
            }
        }

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

        private static void SendMessageToClient(NamedPipeServerStream serverStream, byte[] packet)
        {
            serverStream.Write(packet, 0, packet.Length);
        }

        internal static void SendFilePath(string path)
        {
            byte[] filePathPacket = PackageFilePathMessage(path);
            SendMessageToClient(Server.ServerStream, filePathPacket);
        }
    }
}
