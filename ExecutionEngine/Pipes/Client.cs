//-----------------------------------------------------------------------
// <copyright file="Client.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.IO.Pipes;
using System.Runtime.Serialization.Formatters.Binary;
using VisualStudio.Macros.ExecutionEngine.Pipes;

namespace VSMacros.ExecutionEngine.Pipes
{
    public static class Client
    {
        public static NamedPipeClientStream ClientStream;
        public static void InitializePipeClientStream(Guid guid)
        {
            var timeoutInMilliseconds = 12000;
            Client.ClientStream = new NamedPipeClientStream(".", guid.ToString(), PipeDirection.InOut, PipeOptions.Asynchronous);
            Client.ClientStream.Connect(timeoutInMilliseconds);
        }

        public static void ShutDownServer(NamedPipeClientStream clientStream)
        {
            var packetType = PacketType.Close;
            BinaryFormatter formatter = new BinaryFormatter();
            formatter.Serialize(Client.ClientStream, packetType);
        }

        internal static void SendScriptError(uint modifiedLineNumber, int column, string source, string description)
        {
            var type = PacketType.ScriptError;
            BinaryFormatter formatter = new BinaryFormatter();
            formatter.Serialize(Client.ClientStream, type);

            var scriptError = new ScriptError();
            scriptError.LineNumber = (int)modifiedLineNumber;
            scriptError.Column = column;
            scriptError.Source = source;
            scriptError.Description = description;
            formatter.Serialize(Client.ClientStream, scriptError);
        }

        internal static void SendCriticalError(string message, string source, string stackTrace, string targetSite)
        {
            var type = PacketType.CriticalError;
            BinaryFormatter formatter = new BinaryFormatter();
            formatter.Serialize(Client.ClientStream, type);

            var criticalError = new CriticalError();
            criticalError.Message = message;
            criticalError.Source = source;
            criticalError.StackTrace = stackTrace;
            criticalError.TargetSite = targetSite;
            formatter.Serialize(Client.ClientStream, criticalError);
        }

        internal static void SendSuccessMessage()
        {
            var packetType = PacketType.Success;
            BinaryFormatter formatter = new BinaryFormatter();
            formatter.Serialize(Client.ClientStream, packetType);
        }
    }
}
