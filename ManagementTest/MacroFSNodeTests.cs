//-----------------------------------------------------------------------
// <copyright file="MacroFSNodeTests.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Reflection;
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
            Assert.AreEqual(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Test.js"), this.fileNode.FullPath);
            Assert.IsFalse(this.fileNode.IsEditable);
        }

        [TestMethod]
        public void Directory_Has_IsDirectory_SetToTrue()
        {
            Assert.IsTrue(this.directoryNode.IsDirectory);
        }

        [TestMethod]
        public void File_Has_IsDirectory_SetToFalse()
        {
            Assert.IsFalse(this.fileNode.IsDirectory);
        }

        [TestMethod]
        public void Name_IsFileNameWithoutExtensionFor_FullPath()
        {
            Assert.AreEqual(Path.GetFileNameWithoutExtension(this.fileNode.FullPath), this.fileNode.Name);
        }

        #endregion

        #region Properties

        [TestMethod]
        public void Shortcut_IsBoundToNode()
        {
             // node.Shortcut is initially empty
             Assert.AreEqual(string.Empty, this.fileNode.FormattedShortcut);
            
             // Set the shortcut
             Manager.Shortcuts[1] = this.fileNode.FullPath;
             this.fileNode.Shortcut = MacroFSNode.ToFetch; // notify the change

             // After assigning a shortcut to a node, node.Shortcut should return the formatted shortcut
             string expected = string.Format(MacroFSNode.ShortcutKeys, 1);
             Assert.AreEqual(expected, this.fileNode.FormattedShortcut);
        }

        [TestMethod]
        public void Children_ShouldNotBeEmpty()
        {
            Assert.AreEqual(5, this.directoryNode.Children.Count);
        }

        [TestMethod]
        public void Name_ShouldNotChange_OnInvalidInput()
        {
            string oldName = this.fileNode.Name;

            // Input: Empty string
            this.fileNode.Name = string.Empty;
            Assert.AreEqual(oldName, this.fileNode.Name);

			// Input: Contains invalid chars
            this.fileNode.Name = oldName;
            this.fileNode.Name = "<>:?*";
            Assert.AreEqual(oldName, this.fileNode.Name);

            // Input: Same name
            this.fileNode.Name = oldName;
            this.fileNode.Name = this.fileNode.Name;
            Assert.AreEqual(oldName, this.fileNode.Name);
        }

        #endregion

        #region NotifyPropertyChanged

        // When the name is changed, both the Name and the FullPath properties should be updated
        [Ignore]
        public void Name_NotifiesOnChange()
        {
            List<string> receivedEvents = new List<string>();
            this.fileNode.PropertyChanged += (sender, e) =>
            {
                receivedEvents.Add(e.PropertyName);
            };

            // Change the name of the node
            this.fileNode.Name = "Another Name";

            // Two event are fired
            List<string> expected = new List<string> { "FullPath", "Name" };
            CollectionAssert.AreEquivalent(expected, receivedEvents);
        }

        [TestMethod]
        public void Shortcut_NotifiesOnChange()
        {
            List<string> receivedEvents = new List<string>();
            this.fileNode.PropertyChanged += (sender, e) =>
            {
                receivedEvents.Add(e.PropertyName);
            };

            // Select the node
            this.fileNode.Shortcut = MacroFSNode.ToFetch;

            // Two event are fired
            List<string> expected = new List<string> { "Shortcut", "FormattedShortcut" };
            CollectionAssert.AreEquivalent(expected, receivedEvents);
        }

        [TestMethod]
        public void IsEditable_NotifiesOnChange()
        {
            // By default, the node should not be selected
            Assert.IsFalse(this.fileNode.IsEditable);

            List<string> receivedEvents = new List<string>();
            this.fileNode.PropertyChanged += (sender, e) =>
            {
                receivedEvents.Add(e.PropertyName);
            };

            // Select the node
            this.fileNode.IsEditable = true;

            // Node should now be selected
            Assert.IsTrue(this.fileNode.IsEditable);

            // Only one event should have been fired
            Assert.AreEqual(1, receivedEvents.Count);
            Assert.AreEqual("IsEditable", receivedEvents[0]);
        }

        [TestMethod]
        public void IsSelected_NotifiesOnChange()
        {
            // By default, the node should not be selected
            Assert.IsFalse(this.fileNode.IsSelected);

            List<string> receivedEvents = new List<string>();
            this.fileNode.PropertyChanged += delegate(object sender, PropertyChangedEventArgs e)
            {
                receivedEvents.Add(e.PropertyName);
            };

            // Select the node
            this.fileNode.IsSelected = true;

            // Node should now be selected
            Assert.IsTrue(this.fileNode.IsSelected);
            
            // Only one event should have been fired
            Assert.AreEqual(1, receivedEvents.Count);
            Assert.AreEqual("IsSelected", receivedEvents[0]);
        }

        // When IsEnabled is changed, the class notifies about the change of both the IsExpanded and Icon properties
        [TestMethod]
        public void IsExpanded_NotifiesOnChange()
        {
            List<string> receivedEvents = new List<string>();
            this.fileNode.PropertyChanged += delegate(object sender, PropertyChangedEventArgs e)
            {
                receivedEvents.Add(e.PropertyName);
            };

            // Expand the node
            this.fileNode.IsExpanded = true;

            // Two event are fired
            // Two event are fired
            List<string> expected = new List<string> { "IsExpanded", "Icon" };
            CollectionAssert.AreEquivalent(expected, receivedEvents);
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

            this.fileNode = new MacroFSNode(path);

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
            if (File.Exists(this.fileNode.FullPath))
            {
                File.Delete(this.fileNode.FullPath);
            }
            if (Directory.Exists(this.directoryNode.FullPath))
            {
                Directory.Delete(this.directoryNode.FullPath, true);
            }
        }

        #endregion
    }
}
