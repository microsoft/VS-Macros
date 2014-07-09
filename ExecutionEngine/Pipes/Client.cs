using System;
using System.Collections.Generic;
using System.IO.Pipes;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VSMacros.ExecutionEngine.Helpers;

namespace VSMacros.ExecutionEngine.Pipes
{
    public static class Client
    {
        public static NamedPipeClientStream ClientStream;
        public static void InitializePipeClientStream(string guid)
        {
            //NamedPipeClientStream clientStream = new NamedPipeClientStream(".", guid, PipeDirection.InOut, PipeOptions.Asynchronous);
            Client.ClientStream = new NamedPipeClientStream(".", guid, PipeDirection.InOut, PipeOptions.Asynchronous);
            Client.ClientStream.Connect(120000); // is that unreasonable
            //return clientStream;
        }

        public static void CreateClientPipe(string guid)
        {
            Client.ClientStream = new NamedPipeClientStream(guid);
            Console.WriteLine(string.Format("guid of client string is: {0}", guid));
            Client.ClientStream.Connect(120);
            Console.WriteLine("connected");
        }

        public static void ShutDownServer(NamedPipeClientStream clientStream)
        {
            byte[] close = PackageMessage("close");
            SendMessageToServer(clientStream, close);
        }

        public static void SendMessageToServer(NamedPipeClientStream clientStream, byte[] packet)
        {
            clientStream.Write(packet, 0, packet.Length);
        }

        public static int GetIntFromStream(NamedPipeClientStream clientStream)
        {
            byte[] number = new byte[sizeof(int)];
            clientStream.Read(number, 0, sizeof(int));
            return BitConverter.ToInt32(number, 0);
        }

        public static string GetMessageFromStream(NamedPipeClientStream clientStream, int sizeOfMessage)
        {
            byte[] messageBuffer = new byte[sizeOfMessage];
            clientStream.Read(messageBuffer, 0, sizeOfMessage);
            return UnicodeEncoding.Unicode.GetString(messageBuffer);
        }

        public static byte[] PackageMessage(string line)
        {
            byte[] messageBuffer = UnicodeEncoding.Unicode.GetBytes(line);
            int messageSize = messageBuffer.Length;
            byte[] sizeBuffer = BitConverter.GetBytes(messageSize);

            int intSize = sizeof(int);
            byte[] packet = new byte[intSize + messageSize];
            sizeBuffer.CopyTo(packet, 0);

            int offset = intSize;
            messageBuffer.CopyTo(packet, offset);

            return packet;
        }

        public static string ParseFilePath(NamedPipeClientStream clientStream)
        {
            // Visual Studio -> Execution engine

            int sizeOfMessage = Client.GetIntFromStream(Client.ClientStream);
            string message = Client.GetMessageFromStream(Client.ClientStream, sizeOfMessage);
            return message;
        }

        public static void HandlePacketClose(NamedPipeClientStream clientStream)
        {
            // Visual Studio -> Execution engine
            // Execution engine -> Visual Studio??

            // TODO: Close execution engine from QueryClose in VSMacrosPackage
        }

        public static void HandlePacketSuccess(NamedPipeClientStream clientStream)
        {
            // Execution engine -> Visual Studio
            // TODO: update the CompletedEvent thing
        }

        public static void HandlePacketScriptError(NamedPipeClientStream clientStream)
        {
            // Execution engine -> Visual Studio

            int lineNumber = Client.GetIntFromStream(Client.ClientStream);
            int sizeOfDescription = Client.GetIntFromStream(Client.ClientStream);
            string message = Client.GetMessageFromStream(Client.ClientStream, sizeOfDescription);

            // TODO: update the CompletedEvent thing
        }

        internal static void HandlePacketOtherError(NamedPipeClientStream namedPipeClientStream)
        {
            // Execution engine -> Visual Studio

            int sizeOfDescription = Client.GetIntFromStream(Client.ClientStream);
            string message = Client.GetMessageFromStream(Client.ClientStream, sizeOfDescription);
        }
    }
}
