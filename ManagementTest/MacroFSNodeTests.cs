//-----------------------------------------------------------------------
// <copyright file="MacroFSNodeTests.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VSMacros.Engines;
using VSMacros.Models;

namespace ManagementTest
{
    [TestClass]
    public class MacroFSNodeTests
    {
        #region Initialization

        private MacroFSNode fileNode;
        private MacroFSNode directoryNode;

        [TestMethod]
        public void Constructor_InitializesNode()
        {
            Assert.AreEqual(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Test.js"), fileNode.FullPath);
            Assert.IsFalse(fileNode.IsEditable);
        }

        [TestMethod]
        public void Directory_Has_IsDirectory_SetToTrue()
        {
            Assert.IsTrue(directoryNode.IsDirectory);
        }

        [TestMethod]
        public void File_Has_IsDirectory_SetToFalse()
        {
            Assert.IsFalse(fileNode.IsDirectory);
        }

        [TestMethod]
        public void Name_IsFileNameWithoutExtensionFor_FullPath()
        {
            Assert.AreEqual(Path.GetFileNameWithoutExtension(fileNode.FullPath), fileNode.Name);
        }

        #endregion

        #region Properties

        [Ignore]
        public void Shortcut_IsBoundToNode()
        {
             // node.Shortcut is initially empty
             Assert.AreEqual(string.Empty, fileNode.Shortcut);
            
             // TODO Testing with static instances seems complicated - talk with Ryan about that
             Manager.Instance.Shortcuts[1] = fileNode.FullPath;
            
             // After assigning a shortcut to a node, node.Shortcut should return the formatted shortcut
             Assert.AreEqual("(CTRL+M, 1)", fileNode.Shortcut);
        }

        [TestMethod]
        public void Children_ShouldNotBeEmpty()
        {
            Assert.AreEqual(5, directoryNode.Children.Count);
        }

        #endregion

        #region NotifyPropertyChanged

        // When the name is changed, both the Name and the FullPath properties should be updated
        [Ignore]
        public void Name_NotifiesOnChange()
        {
            List<string> receivedEvents = new List<string>();
            fileNode.PropertyChanged += (sender, e) =>
            {
                receivedEvents.Add(e.PropertyName);
            };

            // Change the name of the node
            fileNode.Name = "Another Name";

            // Two event are fired
            List<string> expected = new List<string> { "FullPath", "Name" };
            CollectionAssert.AreEquivalent(expected, receivedEvents);
        }

        [TestMethod]
        public void Shortcut_NotifiesOnChange()
        {
            string receivedEvent = string.Empty;
            fileNode.PropertyChanged += delegate(object sender, PropertyChangedEventArgs e)
            {
                receivedEvent = e.PropertyName;
            };

            // Select the node
            fileNode.Shortcut = string.Empty;

            // Only one event should have been fired
            Assert.AreEqual("Shortcut", receivedEvent);
        }

        [TestMethod]
        public void IsEditable_NotifiesOnChange()
        {
            // By default, the node should not be selected
            Assert.IsFalse(fileNode.IsEditable);

            List<string> receivedEvents = new List<string>();
            fileNode.PropertyChanged += delegate(object sender, PropertyChangedEventArgs e)
            {
                receivedEvents.Add(e.PropertyName);
            };

            // Select the node
            fileNode.IsEditable = true;

            // Node should now be selected
            Assert.IsTrue(fileNode.IsEditable);

            // Only one event should have been fired
            Assert.AreEqual(1, receivedEvents.Count);
            Assert.AreEqual("IsEditable", receivedEvents[0]);
        }

        [TestMethod]
        public void IsSelected_NotifiesOnChange()
        {
            // By default, the node should not be selected
            Assert.IsFalse(fileNode.IsSelected);

            List<string> receivedEvents = new List<string>();
            fileNode.PropertyChanged += delegate(object sender, PropertyChangedEventArgs e)
            {
                receivedEvents.Add(e.PropertyName);
            };

            // Select the node
            fileNode.IsSelected = true;

            // Node should now be selected
            Assert.IsTrue(fileNode.IsSelected);
            
            // Only one event should have been fired
            Assert.AreEqual(1, receivedEvents.Count);
            Assert.AreEqual("IsSelected", receivedEvents[0]);
        }

        // When IsEnabled is changed, the class notifies about the change of both the IsExpanded and Icon properties
        [TestMethod]
        public void IsExpanded_NotifiesOnChange()
        {
            List<string> receivedEvents = new List<string>();
            fileNode.PropertyChanged += delegate(object sender, PropertyChangedEventArgs e)
            {
                receivedEvents.Add(e.PropertyName);
            };

            // Expand the node
            fileNode.IsExpanded = true;

            // Two event are fired
            Assert.AreEqual(2, receivedEvents.Count);
            Assert.AreEqual("IsExpanded", receivedEvents[0]);
            Assert.AreEqual("Icon", receivedEvents[1]);
        }

        #endregion

        #region Helpers

        [TestInitialize]
        public void CreateFileSystemNodes()
        {
            // Create a file MacroFSNode
            string path = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Test.js");

            if (!File.Exists(path))
            {
                File.Create(path).Close();
            }

            fileNode = new MacroFSNode(path);

            // Create a directory MacroFSNode

            path = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Macros");
            Directory.CreateDirectory(path);

            this.directoryNode = new MacroFSNode(path);

            // Add 5 files to the folder
            for (int i = 0; i < 5; i++)
            {
                path = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Macros", i + ".js");
                File.Create(path).Close();
            }
        }

        [TestCleanup]
        public void CleanUpFileSystem()
        {
            if (File.Exists(fileNode.FullPath))
            {
                File.Delete(fileNode.FullPath);
            }
            if (Directory.Exists(directoryNode.FullPath))
            {
                Directory.Delete(directoryNode.FullPath, true);
            }
        }

        #endregion
    }
}
