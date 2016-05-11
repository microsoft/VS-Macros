//-----------------------------------------------------------------------
// <copyright file="PipeTests.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.IO.Pipes;
using System.Runtime.Serialization.Formatters.Binary;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisualStudio.Macros.ExecutionEngine.Pipes;
using VSMacros.ExecutionEngine.Pipes;
using VSMacros.Pipes;

namespace ExecutionTest
{
    [TestClass]
    class PipeTests
    {
        private NamedPipeServerStream serverStream;
        private NamedPipeClientStream clientStream;
        private BinaryFormatter serializer;

        #region Initialization
        [TestMethod]
        public void Constructor_InitializesPipes()
        {
            Server.InitializeServer();
            var guid = Server.Guid;
            Client.InitializePipeClientStream(guid);

            serverStream = Server.ServerStream;
            clientStream = Client.ClientStream;

            Assert.IsTrue(serverStream.IsConnected);
            Assert.IsTrue(clientStream.IsConnected);
        }

        [TestMethod]
        public void Constructor_InitializesSerializer()
        {
            serializer = new BinaryFormatter();

            Assert.IsTrue(serverStream.IsConnected);
            Assert.IsTrue(clientStream.IsConnected);
        }

        #endregion

        #region Message Passing

        [TestMethod]
        public void ExchangingSuccessMessage_MessageSuccessfullyExchanged()
        {
            var sentPacketType = PacketType.Success;
            serializer.Serialize(clientStream, sentPacketType);
            var receivedType = (PacketType)serializer.Deserialize(serverStream);
            
            Assert.AreEqual(sentPacketType, receivedType);
        }

        [TestMethod]
        public void ExchangingFilePathSimple_FilePathSuccessfullyExchanged()
        {
            var sentFilePath = new FilePath();
            sentFilePath.Iterations = 1;
            sentFilePath.Path = @"C\Windows\file path";

            serializer.Serialize(clientStream, sentFilePath);
            var receivedFilePath = (FilePath)serializer.Deserialize(serverStream);

            Assert.AreEqual(sentFilePath.Iterations, receivedFilePath.Iterations);
            Assert.AreEqual(sentFilePath.Path, receivedFilePath.Path);
        }

        public void ExchangingFilePathWithTypeIndicator_FilePathSuccessfullyExchanged()
        {
            var sentFilePathType = PacketType.FilePath;
            serializer.Serialize(clientStream, sentFilePathType);

            var sentFilePath = new FilePath();
            sentFilePath.Iterations = 1;
            sentFilePath.Path = @"C\Windows\file path";
            serializer.Serialize(clientStream, sentFilePath);

            var receivedFilePathType = (PacketType)serializer.Deserialize(serverStream);
            var receivedFilePath = (FilePath)serializer.Deserialize(serverStream);

            Assert.AreEqual(sentFilePathType, receivedFilePathType);
            Assert.AreEqual(sentFilePath.Iterations, receivedFilePath.Iterations);
            Assert.AreEqual(sentFilePath.Path, receivedFilePath.Path);
        }

        #endregion
    }
}
