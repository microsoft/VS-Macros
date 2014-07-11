//-----------------------------------------------------------------------
// <copyright file="Client.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.IO.Pipes;
using System.Text;
using ExecutionEngine.Enums;

namespace VSMacros.ExecutionEngine.Pipes
{
    public static class Client
    {
        public static NamedPipeClientStream ClientStream;
        public static void InitializePipeClientStream(string guid)
        {
            var someUnreasonableNumber = 1200000;
            Client.ClientStream = new NamedPipeClientStream(".", guid, PipeDirection.InOut, PipeOptions.Asynchronous);
            Client.ClientStream.Connect(someUnreasonableNumber);
        }

        #region Sending
        public static void ShutDownServer(NamedPipeClientStream clientStream)
        {
            byte[] close = PackageCloseMessage();
            SendMessageToServer(clientStream, close);
        }

        public static void SendMessageToServer(NamedPipeClientStream clientStream, byte[] packet)
        {
            clientStream.Write(packet, 0, packet.Length);
        }

        public static byte[] PackageCloseMessage()
        {
            byte[] serializedType = BitConverter.GetBytes((int)Packet.Close);

            int type = sizeof(int);
            byte[] packet = new byte[type];

            serializedType.CopyTo(packet, 0);

            return packet;
        }

        public static byte[] PackageFilePathMessage(string line)
        {
            byte[] serializedType = BitConverter.GetBytes((int)Packet.FilePath);

            byte[] serializedMessage = UnicodeEncoding.Unicode.GetBytes(line);
            int message = serializedMessage.Length;

            byte[] serializedLength = BitConverter.GetBytes(message);

            int type = sizeof(int), messageSize = sizeof(int);
            byte[] packet = new byte[type + messageSize + message];

            serializedType.CopyTo(packet, 0);

            int offset = sizeof(int);
            serializedLength.CopyTo(packet, offset);
            offset += sizeof(int);

            serializedMessage.CopyTo(packet, offset);

            return packet;
        }

        #endregion

        #region Getting
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

        #endregion

        internal static byte[] PackageScriptError(uint lineNumber, string source, string description)
        {
            byte[] serializedType = BitConverter.GetBytes((int)Packet.ScriptError);
            byte[] serializedLineNumber = BitConverter.GetBytes((int)lineNumber);
            byte[] serializedSource = UnicodeEncoding.Unicode.GetBytes(source);
            byte[] serializedDescription = UnicodeEncoding.Unicode.GetBytes(description);
            byte[] serializedSizeOfSource = BitConverter.GetBytes(serializedSource.Length);
            byte[] serializedSizeOfDescription = BitConverter.GetBytes(serializedDescription.Length);

            int typePlaceholder = sizeof(int), lineNumberPlaceholder = sizeof(int), sourceSizePlaceholder = sizeof(int), descriptionSizePlaceholder = sizeof(int);
            int sourcePlaceholder = serializedSource.Length, descriptionPlaceholder = serializedDescription.Length;

            byte[] packet = new byte[typePlaceholder + lineNumberPlaceholder + sourceSizePlaceholder + sourcePlaceholder + descriptionSizePlaceholder + descriptionPlaceholder];

            int offset = 0;
            string type = "type";
            serializedType.CopyTo(packet, offset);

            Console.WriteLine("hello");
            var message = string.Format("at offset {0}, for {1}", offset, type);
            Console.WriteLine(message);

            offset += typePlaceholder;
            type = "line number";
            serializedLineNumber.CopyTo(packet, offset);
            Console.WriteLine(message);

            offset += lineNumberPlaceholder;
            serializedSizeOfSource.CopyTo(packet, offset);
            type = "size of source message";
            Console.WriteLine(message);

            offset += sourceSizePlaceholder;
            serializedSource.CopyTo(packet, offset);
            type = "source mssage";
            Console.WriteLine(message);

            offset += source.Length;
            serializedSizeOfDescription.CopyTo(packet, offset);
            type = "size of description message";
            Console.WriteLine(message);

            offset += descriptionSizePlaceholder;
            serializedDescription.CopyTo(packet, offset);
            type = "description message";
            Console.WriteLine(message);

            for (int i = 0; i < packet.Length; i++)
            {
                Console.Write(packet[i]);
            }
            Console.WriteLine();

            return packet;
        }
    }
}
