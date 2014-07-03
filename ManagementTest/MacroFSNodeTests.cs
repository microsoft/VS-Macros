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

        [TestMethod]
        public void Constructor_InitializesNode()
        {
            MacroFSNode node = this.CreateFileNode();

            Assert.AreEqual(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Macros", "Test.js"), node.FullPath);
            Assert.IsFalse(node.IsEditable);
        }

        [TestMethod]
        public void Directory_Has_IsDirectory_SetToTrue()
        {
            MacroFSNode node = this.CreateDirectoryNode();

            Assert.IsTrue(node.IsDirectory);
        }

        [TestMethod]
        public void File_Has_IsDirectory_SetToFalse()
        {
            MacroFSNode node = this.CreateFileNode();

            Assert.IsFalse(node.IsDirectory);
        }

        [TestMethod]
        public void Name_IsFileNameWithoutExtensionFor_FullPath()
        {
            MacroFSNode node = this.CreateFileNode();

            Assert.AreEqual(Path.GetFileNameWithoutExtension(node.FullPath), node.Name);
        }

        #endregion

        #region Properties

        [TestMethod]
        public void Shortcut_IsBoundToNode()
        {
            MacroFSNode node = this.CreateFileNode();

            // node.Shortcut is initially empty
            // Assert.AreEqual(string.Empty, node.Shortcut);
            //
            // TODO Testing with static instances seems complicated - talk with Ryan about that
            // Manager.Instance.Shortcuts[1] = node.FullPath;
            //
            // After assigning a shortcut to a node, node.Shortcut should return the formatted shortcut
            // Assert.AreEqual("(CTRL+M, 1)", node.Shortcut);
        }

        [TestMethod]
        public void Children_ShouldNotBeEmpty()
        {
            MacroFSNode node = this.CreateDirectoryNode();

            Assert.AreEqual(1, node.Children.Count);
        }

        #endregion

        #region NotifyPropertyChanged

        // When the name is changed, both the Name and the FullPath properties should be updated
        [TestMethod]
        public void Name_NotifiesOnChange()
        {
            MacroFSNode node = this.CreateFileNode();

            List<string> receivedEvents = new List<string>();
            node.PropertyChanged += delegate(object sender, PropertyChangedEventArgs e)
            {
                receivedEvents.Add(e.PropertyName);
            };

            // Change the name of the node
            node.Name = "Test";

            // Two event are fired
            Assert.AreEqual(2, receivedEvents.Count);
            Assert.AreEqual("FullPath", receivedEvents[0]);
            Assert.AreEqual("Name", receivedEvents[1]);
        }

        [TestMethod]
        public void Shortcut_NotifiesOnChange()
        {
            MacroFSNode node = this.CreateFileNode();

            string receivedEvent = string.Empty;
            node.PropertyChanged += delegate(object sender, PropertyChangedEventArgs e)
            {
                receivedEvent = e.PropertyName;
            };

            // Select the node
            node.Shortcut = string.Empty;

            // Only one event should have been fired
            Assert.AreEqual("Shortcut", receivedEvent);
        }

        [TestMethod]
        public void IsEditable_NotifiesOnChange()
        {
            MacroFSNode node = this.CreateFileNode();

            // By default, the node should not be selected
            Assert.IsFalse(node.IsEditable);

            List<string> receivedEvents = new List<string>();
            node.PropertyChanged += delegate(object sender, PropertyChangedEventArgs e)
            {
                receivedEvents.Add(e.PropertyName);
            };

            // Select the node
            node.IsEditable = true;

            // Node should now be selected
            Assert.IsTrue(node.IsEditable);

            // Only one event should have been fired
            Assert.AreEqual(1, receivedEvents.Count);
            Assert.AreEqual("IsEditable", receivedEvents[0]);
        }

        [TestMethod]
        public void IsSelected_NotifiesOnChange()
        {
            MacroFSNode node = this.CreateFileNode();

            // By default, the node should not be selected
            Assert.IsFalse(node.IsSelected);

            List<string> receivedEvents = new List<string>();
            node.PropertyChanged += delegate(object sender, PropertyChangedEventArgs e)
            {
                receivedEvents.Add(e.PropertyName);
            };

            // Select the node
            node.IsSelected = true;

            // Node should now be selected
            Assert.IsTrue(node.IsSelected);
            
            // Only one event should have been fired
            Assert.AreEqual(1, receivedEvents.Count);
            Assert.AreEqual("IsSelected", receivedEvents[0]);
        }

        // When IsEnabled is changed, the class notifies about the change of both the IsExpanded and Icon properties
        [TestMethod]
        public void IsExpanded_NotifiesOnChange()
        {
            MacroFSNode node = this.CreateFileNode();

            List<string> receivedEvents = new List<string>();
            node.PropertyChanged += delegate(object sender, PropertyChangedEventArgs e)
            {
                receivedEvents.Add(e.PropertyName);
            };

            // Expand the node
            node.IsExpanded = true;

            // Two event are fired
            Assert.AreEqual(2, receivedEvents.Count);
            Assert.AreEqual("IsExpanded", receivedEvents[0]);
            Assert.AreEqual("Icon", receivedEvents[1]);
        }

        #endregion

        #region Helpers

        private MacroFSNode CreateFileNode()
        {
            string path = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Macros", "Test.js");
            return new MacroFSNode(path);
        }

        private MacroFSNode CreateDirectoryNode()
        {
            string path = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Macros");
            return new MacroFSNode(path);
        }

        #endregion
    }
}
