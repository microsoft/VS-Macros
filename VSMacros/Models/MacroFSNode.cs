using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.IO;
using System.Windows;
using GelUtilities = Microsoft.Internal.VisualStudio.PlatformUI.Utilities;
using System.Windows.Media.Imaging;

namespace VSMacros.Models
{
    public sealed class MacroFSNode : INotifyPropertyChanged
    {
        private string fullPath;
        private bool isEditable;
        private bool isExpanded;

        private MacroFSNode parent;
        private ObservableCollection<MacroFSNode> children;
        static private HashSet<string> enabledDirectories = new HashSet<string>();

        public event PropertyChangedEventHandler PropertyChanged;

        public static MacroFSNode RootNode;

        public bool IsDirectory { get; private set; }

        public MacroFSNode(string path, MacroFSNode parent = null)
        {
            this.IsDirectory = (File.GetAttributes(path) & FileAttributes.Directory) == FileAttributes.Directory;
            this.FullPath = path;
            this.isEditable = false;
            this.parent = parent;
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
                string oldFullPath = this.FullPath;
                string newFullPath = Path.Combine(Path.GetDirectoryName(this.FullPath), value + Path.GetExtension(this.FullPath));

                try
                {
                     // Update file system
                    if (this.IsDirectory)
                    {
                        Directory.Move(oldFullPath, newFullPath);
                    }
                    else
                    {
                        File.Move(oldFullPath, newFullPath);
                    }

                    // Update object
                    this.FullPath = newFullPath;
                }
                catch(Exception e)
                {
                    if (e.Message != null)
                    {
                        MessageBox.Show(e.Message);
                    }
                }
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

        public string Icon
        {
            get
            {
                //IVsImageService imageService = (IVsImageService)((IServiceProvider)VSMacrosPackage.Current).GetService(typeof(SVsImageService));

                //if (imageService != null)
                //{
                //    IVsUIObject uiObject = imageService.GetIconForFile(Path.GetFileName(this.FullPath), __VSUIDATAFORMAT.VSDF_WPF);

                //    if (uiObject != null)
                //    {
                //        BitmapSource bitmapSource = GelUtilities.GetObjectData(uiObject) as BitmapSource;

                //        return bitmapSource.ToString();
                //    }
                //}

                //return string.Empty;

                if (this.IsDirectory)
                {
                    if (this.isExpanded)
                    {
                        return @"..\Resources\folderopened.png";
                    }
                    else
                    {
                        return @"..\Resources\folderclosed.png";
                    }
                }
                else
                {
                    return @"..\Resources\js.png";
                }
            }
        }

        public bool IsExpanded
        {
            get
            {
                return this.isExpanded;
            }

            set
            {
                this.isExpanded = value;

                if (this.IsExpanded)
                {
                    enabledDirectories.Add(this.FullPath);
                }
                else
                {
                    enabledDirectories.Remove(this.FullPath);
                }

                NotifyPropertyChanged("IsExpanded");
                NotifyPropertyChanged("Icon");
            }
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
                    this.children = this.GetChildNodes();
                }

                return this.children;
            }
        }

        // Returns a list of children of the current node
        private ObservableCollection<MacroFSNode> GetChildNodes()
        {
            var files = from childFile in Directory.GetFiles(this.FullPath)
                       where Path.GetExtension(childFile) == ".js"
                       orderby childFile
                       select childFile;

            var directories = from childDirectory in Directory.GetDirectories(this.FullPath)
                              orderby childDirectory
                              select childDirectory;

            return new ObservableCollection<MacroFSNode>(files.Union(directories)
                    .Select((item) => new MacroFSNode(item, this)));
        }

        #region Context Menu
        public void Delete()
        {
            this.parent.children.Remove(this);
        }

        public void EnableEdit()
        {
            this.IsEditable = true;
        }

        public void DisableEdit()
        {
            this.IsEditable = false;
        }

        public void RefreshTree()
        {
            // Make a copy of the hashset
            HashSet<string> dirs = new HashSet<string>(enabledDirectories);

            // Clear enableDirectories
            enabledDirectories.Clear();

            // Refetch the children of the root node
            RootNode.children = this.GetChildNodes();

            // Recursively set IsEnabled for each folders
            SetIsExpanded(RootNode, dirs);

            // Notify change
            NotifyPropertyChanged("Children");
        }
        #endregion
        
        // OPTIMIZATION IDEA instead of iterating over the children, iterate over the enableDirs
        private void SetIsExpanded(MacroFSNode node, HashSet<string> enabledDirs)
        {
            if (node.Children.Count > 0 && enabledDirs.Count > 0)
            {
                foreach (var item in node.children)
                {
                    if (item.IsDirectory && enabledDirs.Contains(item.FullPath))
                    {
                        // Set IsExpanded
                        item.IsExpanded = true;

                        // Remove path from dirs
                        enabledDirs.Remove(item.FullPath);

                        // Recursion on children
                        SetIsExpanded(item, enabledDirs);
                    }
                }
            }
        }

        // Create the OnPropertyChanged method to raise the event 
        private void NotifyPropertyChanged(string name)
        {
            PropertyChangedEventHandler handler = this.PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(name));
            }
        }
    }
}
