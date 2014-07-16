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
        public static void InitializePipeClientStream(Guid guid)
        {
            var timeoutInMilliseconds = 120000;
            Client.ClientStream = new NamedPipeClientStream(".", guid.ToString(), PipeDirection.InOut, PipeOptions.Asynchronous);
            Client.ClientStream.Connect(timeoutInMilliseconds);
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

        internal static byte[] PackageCriticalError(string message, string source, string stackTrace, string targetSite)
        {
            byte[] serializedType = BitConverter.GetBytes((int)Packet.CriticalError);

            byte[] serializedMessage = UnicodeEncoding.Unicode.GetBytes(message);
            byte[] serializedSource = UnicodeEncoding.Unicode.GetBytes(source);
            byte[] serializedStackTrace = UnicodeEncoding.Unicode.GetBytes(stackTrace);
            byte[] serializedTargetSite = UnicodeEncoding.Unicode.GetBytes(targetSite);

            byte[] serializedSizeOfMessage = BitConverter.GetBytes(serializedMessage.Length);
            byte[] serializedSizeOfSource = BitConverter.GetBytes(serializedSource.Length);
            byte[] serializedSizeOfStackTrace = BitConverter.GetBytes(serializedStackTrace.Length);
            byte[] serializedSizeOfTargetSite = BitConverter.GetBytes(serializedTargetSite.Length);

            int type = sizeof(int), messageSize = sizeof(int), sourceSize = sizeof(int), stackTraceSize = sizeof(int), targetSiteSize = sizeof(int);

            byte[] packet = new byte[type + messageSize + serializedMessage.Length + sourceSize + serializedSource.Length + 
                stackTraceSize + serializedStackTrace.Length + targetSiteSize + serializedTargetSite.Length];

            int offset = 0;
            serializedType.CopyTo(packet, offset);

            offset += serializedType.Length;
            serializedSizeOfMessage.CopyTo(packet, offset);

            offset += serializedSizeOfMessage.Length;
            serializedMessage.CopyTo(packet, offset);

            offset += serializedMessage.Length;
            serializedSizeOfSource.CopyTo(packet, offset);

            offset += serializedSizeOfSource.Length;
            serializedSource.CopyTo(packet, offset);

            offset += serializedSource.Length;
            serializedSizeOfStackTrace.CopyTo(packet, offset);

            offset += serializedSizeOfStackTrace.Length;
            serializedStackTrace.CopyTo(packet, offset);

            offset += serializedStackTrace.Length;
            serializedSizeOfTargetSite.CopyTo(packet, offset);

            offset += serializedSizeOfTargetSite.Length;
            serializedTargetSite.CopyTo(packet, offset);

            return packet;
        }

        internal static byte[] PackageScriptError(uint errorLineNumber, int errorColumn, string errorSource, string errorDescription)
        {
            byte[] serializedType = BitConverter.GetBytes((int)Packet.ScriptError);
            byte[] serializedLineNumber = BitConverter.GetBytes((int)errorLineNumber);
            byte[] serializedCharacterPos = BitConverter.GetBytes((int)errorColumn);
            byte[] serializedSource = UnicodeEncoding.Unicode.GetBytes(errorSource);
            byte[] serializedDescription = UnicodeEncoding.Unicode.GetBytes(errorDescription);
            byte[] serializedSizeOfSource = BitConverter.GetBytes(serializedSource.Length);
            byte[] serializedSizeOfDescription = BitConverter.GetBytes(serializedDescription.Length);

            int type = sizeof(int);
            int lineNumber = sizeof(int);
            int column = sizeof(int), sourceSize = sizeof(int), descriptionSize = sizeof(int);
            int source = serializedSource.Length, description = serializedDescription.Length;

            byte[] packet = new byte[type + lineNumber + column + sourceSize + source + descriptionSize + description];

            int offset = 0;
            serializedType.CopyTo(packet, offset);

            offset += serializedType.Length;
            serializedLineNumber.CopyTo(packet, offset);

            offset += serializedLineNumber.Length;
            serializedCharacterPos.CopyTo(packet, offset);

            offset += serializedCharacterPos.Length;
            serializedSizeOfSource.CopyTo(packet, offset);

            offset += serializedSizeOfSource.Length;
            serializedSource.CopyTo(packet, offset);

            offset += serializedSource.Length;
            serializedSizeOfDescription.CopyTo(packet, offset);

            offset += serializedSizeOfDescription.Length;
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

        public static string GetFilePath(NamedPipeClientStream clientStream)
        {
            int sizeOfMessage = Client.GetInt(Client.ClientStream);
            string message = Client.GetMessage(Client.ClientStream, sizeOfMessage);
            return message;
        }

        #endregion
    }
}
