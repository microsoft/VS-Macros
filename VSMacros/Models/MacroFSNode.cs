using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.IO;
using System.Windows;

namespace VSMacros.Model
{
    public sealed class MacroFSNode : INotifyPropertyChanged
    {
        private string fullPath;
        private bool isEditable;

        private MacroFSNode parent;
        private ObservableCollection<MacroFSNode> children;

        public event PropertyChangedEventHandler PropertyChanged;
        public bool IsDirectory { get; private set; }

        public MacroFSNode(string path, MacroFSNode parent = null)
        {
            this.FullPath = path;
            this.IsDirectory = (File.GetAttributes(this.FullPath) & FileAttributes.Directory) == FileAttributes.Directory;
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
                if (this.IsDirectory)
                {
                    return @"..\Resources\folder.png";
                }
                else
                {
                    return @"..\Resources\js.png";
                }
            }
        }

        public IEnumerable<MacroFSNode> Children
        {
            get
            {
                if (!this.IsDirectory)
                {
                    return Enumerable.Empty<MacroFSNode>();
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
            var file = from childFile in Directory.GetFiles(this.FullPath)
                       where Path.GetExtension(childFile) == ".js"
                       orderby childFile
                       select childFile;

            var directories = from childDirectory in Directory.GetDirectories(this.FullPath)
                              orderby childDirectory
                              select childDirectory;

            return new ObservableCollection<MacroFSNode>(directories.Union(file)
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
            this.isEditable = false;
        }

        public void Refresh()
        {
        }

        #endregion

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
