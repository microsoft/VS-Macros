//-----------------------------------------------------------------------
// <copyright file="MacroFSNode.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Media.Imaging;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using VSMacros.Engines;
using GelUtilities = Microsoft.Internal.VisualStudio.PlatformUI.Utilities;

namespace VSMacros.Models
{
    public sealed class MacroFSNode : INotifyPropertyChanged
    {
        // HashSet containing the enabled directories
        private static HashSet<string> enabledDirectories = new HashSet<string>();
        public static HashSet<string> EnabledDirectories { 
            get 
            { 
                return MacroFSNode.enabledDirectories; 
            } 

            set
            {
                MacroFSNode.enabledDirectories = value;
                MacroFSNode.RootNode.SetIsExpanded(MacroFSNode.RootNode, MacroFSNode.enabledDirectories);
            }
        }

        // Properties the binding client watches
        private string fullPath;
        private int shortcut;
        private bool isEditable;
        private bool isExpanded;
        private bool isSelected;
        private bool isMatch;

        private readonly MacroFSNode parent;
        private ObservableCollection<MacroFSNode> children;

        // Constants
        public const int ToFetch = -1;
        public const int None = 0;
        public const string ShortcutKeys = "(CTRL+M, {0})";

        // Static members
        public static MacroFSNode RootNode { get; set; }
        private static bool Searching = false;

        // For INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;

        public MacroFSNode(string path, MacroFSNode parent = null)
        {
            this.IsDirectory = (File.GetAttributes(path) & FileAttributes.Directory) == FileAttributes.Directory;
            this.FullPath = path;
            this.shortcut = MacroFSNode.ToFetch;
            this.isEditable = false;
            this.isSelected = false;
            this.isExpanded = false;
            this.isMatch = false;
            this.parent = parent;

            // Monitor that node 
            //FileChangeMonitor.Instance.MonitorFileSystemEntry(this.FullPath, this.IsDirectory);
        }

        public string FullPath
        {
            get
            { 
                return this.fullPath;
            }

            set 
            { 
                this.fullPath = value;
                this.NotifyPropertyChanged("FullPath");
                this.NotifyPropertyChanged("Name");
            }
        }

        public string Name
        {
            get
            {
                string path = Path.GetFileNameWithoutExtension(this.FullPath);

                if (string.IsNullOrWhiteSpace(path))
                {
                    return this.FullPath;
                }

                return path;
            }

            set
            {
                try
                {
                    // Path.GetFullPath will throw an exception if the path is invalid
                    Path.GetFileName(value);

                    if (value != this.Name && !string.IsNullOrWhiteSpace(value))
                    {
                        string oldFullPath = this.FullPath;
                        string newFullPath = Path.Combine(Path.GetDirectoryName(this.FullPath), value + Path.GetExtension(this.FullPath));

                        // Update file system
                        if (this.IsDirectory)
                        {
                            Directory.Move(oldFullPath, newFullPath);

                            if (MacroFSNode.enabledDirectories.Remove(oldFullPath))
                            {
                                MacroFSNode.enabledDirectories.Add(newFullPath);
                            }
                        }
                        else
                        {
                            File.Move(oldFullPath, newFullPath);
                        }

                        // Update object
                        this.FullPath = newFullPath;

                        // Update shortcut
                        if (this.Shortcut >= MacroFSNode.None)
                        {
                            Manager.Shortcuts[this.shortcut] = newFullPath;
                            Manager.Instance.SaveShortcuts(true);
                        }
                    }
                }
                catch (Exception e)
                {
                    if (e.Message != null)
                    {
                        Manager.Instance.ShowMessageBox(e.Message);
                        MacroFSNode.RefreshTree();
                    }
                }
            }
        }

        public int Shortcut
        {
            get
            {
                return this.shortcut;
            }

            set
            {
                // Shortcut will be refetched
                this.shortcut = MacroFSNode.ToFetch;

                // Just notify the binding
                this.NotifyPropertyChanged("Shortcut");
                this.NotifyPropertyChanged("FormattedShortcut");
            }
        }

        public string FormattedShortcut
        {
            get
            {
                if (this.shortcut == MacroFSNode.ToFetch)
                {
                    
                    this.shortcut = MacroFSNode.None;

                    // Find shortcut, if it exists
                    for (int i = 1; i < 10; i++)
                    {
                        if (string.Compare(Manager.Shortcuts[i], this.FullPath, StringComparison.OrdinalIgnoreCase) == 0)
                        {
                            this.shortcut = i;
                        }
                    }                  
                }

                if (this.shortcut != MacroFSNode.None)
                {
                    return string.Format(MacroFSNode.ShortcutKeys, this.shortcut);
                }

                return string.Empty;
            }
        }

        public BitmapSource Icon
        {
            get
            {
                if (this.IsDirectory)
                {
                    Bitmap bmp;

                    if (this == MacroFSNode.RootNode)
                    {
                        bmp = Resources.RootIcon;
                    }
                    else if (this.isExpanded)
                    {
                        bmp = Resources.FolderOpenedIcon;
                    }
                    else
                    {
                        bmp = Resources.FolderClosedIcon;
                    }

                    return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                            bmp.GetHbitmap(),
                            IntPtr.Zero,
                            Int32Rect.Empty,
                            BitmapSizeOptions.FromEmptyOptions());
                }
                else
                {
                    IVsImageService imageService = (IVsImageService)((IServiceProvider)VSMacrosPackage.Current).GetService(typeof(SVsImageService));
                    if (imageService != null)
                    {
                        IVsUIObject uiObject = imageService.GetIconForFile(Path.GetFileName(this.FullPath), __VSUIDATAFORMAT.VSDF_WPF);
                        if (uiObject != null)
                        {
                            BitmapSource bitmapSource = GelUtilities.GetObjectData(uiObject) as BitmapSource;
                            return bitmapSource;
                        }
                    }
                }

                // Would it be better to have a sane default image?
                return null;
            }
        }

        public bool IsEditable
        {
            get 
            { 
                return this.isEditable; 
            }

            set
            {
                this.isEditable = value;
                this.NotifyPropertyChanged("IsEditable");
            }
        }

        public bool IsSelected
        {
            get
            {
                return this.isSelected;
            }

            set
            {
                this.isSelected = value;
                this.NotifyPropertyChanged("IsSelected");
            }
        }

        public bool IsExpanded
        {
            get
            {
                return this.isExpanded || this.isMatch;
            }

            set
            {
                if (this.IsDirectory)
                {
                    this.isExpanded = value;

                    if (this.IsExpanded)
                    {

                            MacroFSNode.enabledDirectories.Add(this.FullPath);

                        // Expand parent as well
                        if (this.parent != null)
                        {
                            this.parent.IsExpanded = true;
                        }
                    }
                    else
                    {

                            MacroFSNode.enabledDirectories.Remove(this.FullPath);
                    }

                    this.NotifyPropertyChanged("IsExpanded");
                    this.NotifyPropertyChanged("Icon");
                }
            }
        }

        public bool IsMatch
        {
            get
            {
                // If searching is not enabled, always return true
                if (!MacroFSNode.Searching)
                {
                    return true;
                }
                else
                {
                    return this.isMatch;
                }
            }
            set
            {
                this.isMatch = value;

                if (this.IsMatch && this.parent != null)
                {
                    //this.isExpanded = true;
                    this.parent.IsMatch = true;
                }

                this.NotifyPropertyChanged("IsExpanded");
                this.NotifyPropertyChanged("IsMatch");
            }
        }

        public bool IsNotRoot
        {
            get
            {
                return this != MacroFSNode.RootNode;
            }
        }

        public int Depth
        {
            get
            {
                if (this.parent == null)
                {
                    return 0;
                }
                else
                {
                    return this.parent.Depth + 1;
                }
            }
        }

        public bool IsDirectory { get; private set; }

        public MacroFSNode Parent
        {
            get
            {
                return this.parent == null ? MacroFSNode.RootNode : this.parent;
            }
        }

        public bool Equals(MacroFSNode node)
        {
            return this.FullPath == node.FullPath;
        }

        public ObservableCollection<MacroFSNode> Children
        {
            get
            {
                if (!this.IsDirectory)
                {
                    return null;
                }

                if (this.children == null)
                {
                    // Initialize children
                    //this.children = new ObservableCollection<MacroFSNode>();

                    this.children = this.GetChildNodes();

                    // Retrieve children in a background thread
                    // TODO problem: SetIsExpanded goes down the tree and enables previously enabled folders -> if fetching the child is done in the background, the method thinks that the folder has no child and skips it
                    //Task.Run(() => { this.children = this.GetChildNodes(); })
                    //   .ContinueWith(_ => this.NotifyPropertyChanged("Children"), TaskScheduler.FromCurrentSynchronizationContext());
                }

                return this.children;
            }
        }

        private ObservableCollection<MacroFSNode> GetChildNodes()
        {
            var files = from childFile in Directory.GetFiles(this.FullPath)
                       where Path.GetExtension(childFile) == ".js"
                       where childFile != Manager.CurrentMacroPath
                       orderby childFile
                       select childFile;

            var directories = from childDirectory in Directory.GetDirectories(this.FullPath)
                              orderby childDirectory
                              select childDirectory;

            // Merge files and directories into a collection
            ObservableCollection<MacroFSNode> collection = 
                new ObservableCollection<MacroFSNode>(
                    files.Union(directories)
                         .Select((item) => new MacroFSNode(item, this)));

            // Add Current macro at the beginning if this is the root node
            if (this == MacroFSNode.RootNode)
            {
                collection.Insert(0, new MacroFSNode(Manager.CurrentMacroPath, this));
            }

            return collection;
        }

        public void Delete()
        {
            // If a shortcut is bound to the macro
            if (this.shortcut > 0)
            {
                // Remove shortcut from shortcut list
                Manager.Shortcuts[this.shortcut] = string.Empty;
                Manager.Instance.SaveShortcuts(true);
            }

            // Remove macro from collection
            this.parent.children.Remove(this);

            // Unmonitor the file
            //FileChangeMonitor.Instance.UnmonitorFileSystemEntry(this.FullPath, this.IsDirectory);
        }

        public void EnableEdit()
        {
            this.IsEditable = true;
        }

        public void DisableEdit()
        {
            this.IsEditable = false;
        }

        public static void EnableSearch()
        {
            // Set Searching to true
            MacroFSNode.Searching = true;

            // And then notify all node that their IsMatch property might have changed
            MacroFSNode.NotifyAllNode(MacroFSNode.RootNode, "IsMatch");
        }

        public static void DisableSearch()
        {
            // Set Searching to true
            MacroFSNode.Searching = false;

            MacroFSNode.UnmatchAllNodes(MacroFSNode.RootNode);
        }

        /// <summary>
        /// Finds the node with FullPath path in the entire tree 
        /// </summary>
        /// <param name="path"></param>
        /// <returns>MacroFSNode  whose FullPath is path</returns>
        public static MacroFSNode FindNodeFromFullPath(string path)
        {
            if (MacroFSNode.RootNode == null)
            {
                return null;
            }

            // Default node if search fails
            MacroFSNode defaultNode = MacroFSNode.RootNode.Children.Count > 0 ? MacroFSNode.RootNode.Children[0] : MacroFSNode.RootNode;

            // Make sure path is a valid string
            if (string.IsNullOrEmpty(path))
            {
                return defaultNode;
            }

            // Split the string at '\'
            string shortenPath = path.Substring(path.IndexOf(@"\Macros"));
            string[] substrings = shortenPath.Split(new char[] { '\\' });

            // Starting from the root,
            MacroFSNode node = MacroFSNode.RootNode;

            try
            {
                // Go down the tree to find the right node
                // 2 because substrings[0] == "" and substrings[1] is root
                for (int i = 3; i < substrings.Length; i++)
                {
                    node = node.Children.Single(x => x.Name == Path.GetFileNameWithoutExtension(substrings[i]));
                }
            }
            catch (Exception e)
            {
                if (ErrorHandler.IsCriticalException(e))
                {
                    throw;
                }

                // Return default node
                node = defaultNode;
            }

            return node;
        }

        public static MacroFSNode SelectNode(string path)
        {
            // Find node
            MacroFSNode node = FindNodeFromFullPath(path);
            if (node != null)
            {
                // Select it
                node.IsSelected = true;
            }

            return node;
        }

        public static void RefreshTree()
        {
            MacroFSNode root = MacroFSNode.RootNode;
            MacroFSNode.RefreshTree(root);
        }

        private void AfterRefresh(MacroFSNode root, string selectedPath, HashSet<string> dirs)
        {
            // Set IsEnabled for each folders
            root.SetIsExpanded(root, dirs);

            // Selecte the previously selected macro
            MacroFSNode selected = MacroFSNode.FindNodeFromFullPath(selectedPath);
            selected.IsSelected = true;

            // Notify change
            root.NotifyPropertyChanged("Children");
        }

        public static void RefreshTree(MacroFSNode root)
        {
            MacroFSNode selected = MacrosControl.Current.MacroTreeView.SelectedItem as MacroFSNode;

            // Make a copy of the hashset
            HashSet<string> dirs = new HashSet<string>(enabledDirectories);

            // Clear enableDirectories
            enabledDirectories.Clear();

            // Retrieve children in a background thread
            //Task.Run(() => root.children = root.GetChildNodes())
            //    .ContinueWith(_ => root.AfterRefresh(root, selected.FullPath, dirs), TaskScheduler.FromCurrentSynchronizationContext());
            root.children = root.GetChildNodes();
            root.AfterRefresh(root, selected.FullPath, dirs);
        }

        public static void CollapseAllNodes(MacroFSNode root)
        {
            if (root.Children != null)
            {
                foreach (var child in root.Children)
                {
                    child.IsExpanded = false;
                    MacroFSNode.CollapseAllNodes(child);
                }
            }
        }

        public static void UnmatchAllNodes(MacroFSNode root)
        {
            root.isMatch = false;
            root.NotifyPropertyChanged("IsMatch");
            root.NotifyPropertyChanged("IsExpanded");

            if (root.Children != null)
            {
                foreach (var child in root.Children)
                {
                    MacroFSNode.UnmatchAllNodes(child);
                }
            }
        }

        /// <summary>
        /// Expands all the node marked as expanded in <paramref name="enabledDirs"/>.
        /// </summary>
        /// <param name="node">Tree rooted at node.</param>
        /// <param name="enabledDirs">Hash set containing the enabled dirs.</param>
        private void SetIsExpanded(MacroFSNode node, HashSet<string> enabledDirs)
        {
            node.IsExpanded = true;

            // OPTIMIZATION IDEA instead of iterating over the children, iterate over the enableDirs
            if (node.Children.Count > 0 && enabledDirs.Count > 0)
            {
                foreach (var item in node.children)
                {
                    if (item.IsDirectory && enabledDirs.Remove(item.FullPath))
                    {
                        // Set IsExpanded
                        item.IsExpanded = true;

                        // Recursion on children
                        this.SetIsExpanded(item, enabledDirs);
                    }
                }
            }
        }

        private void NotifyPropertyChanged(string name)
        {
            PropertyChangedEventHandler handler = this.PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(name));
            }
        }

        // Notifies all the nodes of the tree rooted at 'node'
        public static void NotifyAllNode(MacroFSNode root, string property)
        {
            root.NotifyPropertyChanged(property);

            if (root.Children != null)
            {
                foreach (var child in root.Children)
                {
                    MacroFSNode.NotifyAllNode(child, property);
                }
            }
        }
    }
}