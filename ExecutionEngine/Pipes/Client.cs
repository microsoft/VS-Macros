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
            var someBigNumber = 1200000;
            Client.ClientStream = new NamedPipeClientStream(".", guid, PipeDirection.InOut, PipeOptions.Asynchronous);
            Client.ClientStream.Connect(someBigNumber);
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

        internal static byte[] PackageScriptError(uint errorLineNumber, int errorCharacterPos, string errorSource, string errorDescription)
        {
            byte[] serializedType = BitConverter.GetBytes((int)Packet.ScriptError);
            byte[] serializedLineNumber = BitConverter.GetBytes((int)errorLineNumber);
            byte[] serializedCharacterPos = BitConverter.GetBytes((int)errorCharacterPos);
            byte[] serializedSource = UnicodeEncoding.Unicode.GetBytes(errorSource);
            byte[] serializedDescription = UnicodeEncoding.Unicode.GetBytes(errorDescription);
            byte[] serializedSizeOfSource = BitConverter.GetBytes(serializedSource.Length);
            byte[] serializedSizeOfDescription = BitConverter.GetBytes(serializedDescription.Length);

            int type = sizeof(int), lineNumber = sizeof(int), characterPos = sizeof(int), sourceSize = sizeof(int), descriptionSize = sizeof(int);
            int source = serializedSource.Length, description = serializedDescription.Length;

            byte[] packet = new byte[type + lineNumber + characterPos + sourceSize + source + descriptionSize + description];

            int offset = 0;
            serializedType.CopyTo(packet, offset);

            offset += type;
            serializedLineNumber.CopyTo(packet, offset);

            offset += lineNumber;
            serializedCharacterPos.CopyTo(packet, offset);

            offset += characterPos;
            serializedSizeOfSource.CopyTo(packet, offset);

            offset += sourceSize;
            serializedSource.CopyTo(packet, offset);

            offset += serializedSource.Length;
            serializedSizeOfDescription.CopyTo(packet, offset);

            offset += descriptionSize;
            serializedDescription.CopyTo(packet, offset);

            return packet;
        }

        internal static byte[] PackageSuccessMessage()
        {
            return BitConverter.GetBytes((int)Packet.Success);
        }

        #endregion

        #region Getting

        internal static int GetIterations(NamedPipeClientStream clientStream)
        {
            int iterations = Client.GetInt(Client.ClientStream);
            return iterations;
        }

        public static int GetInt(NamedPipeClientStream clientStream)
        {
            byte[] number = new byte[sizeof(int)];
            clientStream.Read(number, 0, sizeof(int));

            var intFromStream = BitConverter.ToInt32(number, 0);
            return intFromStream;
        }

        public static string GetMessage(NamedPipeClientStream clientStream, int sizeOfMessage)
        {
            byte[] messageBuffer = new byte[sizeOfMessage];
            clientStream.Read(messageBuffer, 0, sizeOfMessage);
            return UnicodeEncoding.Unicode.GetString(messageBuffer);
        }

        public static string ParseFilePath(NamedPipeClientStream clientStream)
        {
            int sizeOfMessage = Client.GetInt(Client.ClientStream);
            string message = Client.GetMessage(Client.ClientStream, sizeOfMessage);
            return message;
        }

        #endregion
    }
}
