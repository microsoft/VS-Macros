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
        private static BinaryFormatter serializer;

        public static void InitializePipeClientStream(Guid guid)
        {
            var timeoutInMilliseconds = 12000;
            Client.ClientStream = new NamedPipeClientStream(".", guid.ToString(), PipeDirection.InOut, PipeOptions.Asynchronous);
            Client.ClientStream.Connect(timeoutInMilliseconds);

            Client.serializer = new BinaryFormatter();
            Client.serializer.Binder = new BinderHelper();
        }

        public static void ShutDownServer(NamedPipeClientStream clientStream)
        {
            var packetType = PacketType.Close;
            Client.serializer.Serialize(Client.ClientStream, packetType);
        }

        internal static void SendGenericScriptError(uint modifiedLineNumber, int column, string source, string description)
        {
            var type = PacketType.GenericScriptError;
            Client.serializer.Serialize(Client.ClientStream, type);

            var scriptError = new GenericScriptError();
            scriptError.LineNumber = (int)modifiedLineNumber;
            scriptError.Column = column;
            scriptError.Source = source;
            scriptError.Description = description;
            Client.serializer.Serialize(Client.ClientStream, scriptError);
        }

        internal static void SendCriticalError(string message, string source, string stackTrace, string targetSite)
        {
            var type = PacketType.CriticalError;
            BinaryFormatter formatter = new BinaryFormatter();
            Client.serializer.Serialize(Client.ClientStream, type);

            var criticalError = new CriticalError();
            criticalError.Message = message;
            criticalError.Source = source;
            criticalError.StackTrace = stackTrace;
            criticalError.TargetSite = targetSite;
            Client.serializer.Serialize(Client.ClientStream, criticalError);
        }

        internal static void SendSuccessMessage()
        {
            var packetType = PacketType.Success;
            Client.serializer.Serialize(Client.ClientStream, packetType);
        }
    }
}
