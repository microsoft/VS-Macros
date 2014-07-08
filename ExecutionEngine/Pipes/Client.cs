using System;
using System.Collections.Generic;
using System.IO.Pipes;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

        public static int GetSizeOfMessageFromStream(NamedPipeClientStream clientStream)
        {
            int intSize = sizeof(int);
            byte[] sizeBuffer = new byte[intSize];
            clientStream.Read(sizeBuffer, 0, intSize);
            return BitConverter.ToInt32(sizeBuffer, 0);
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
    }
}
