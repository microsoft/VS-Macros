using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.IO;
using VSMacros.Model;

namespace VSMacros
{
    public partial class MacrosControl : UserControl
    {
        private MacroFSNode RootNode;

        public MacrosControl(MacroFSNode rootNode)
        {
            this.RootNode = rootNode;

            // Let the UI bind to the view-model
            this.DataContext = new MacroFSNode[] { this.RootNode };
            InitializeComponent();
        }

        public string SelectedPath { get; set; }

        #region Context Menu Handlers
        private void Playback(object sender, RoutedEventArgs e)
        {

        }

        private void Edit(object sender, RoutedEventArgs e)
        {
            MacroFSNode item = macroTreeView.SelectedItem as MacroFSNode;
            string path = item.FullPath;

            EnvDTE.DTE dte = (EnvDTE.DTE)System.Runtime.InteropServices.Marshal.GetActiveObject("VisualStudio.DTE.12.0");
            dte.ItemOperations.OpenFile(path);
        }

        // TODO Refactor some more
        private void Delete(object sender, RoutedEventArgs e)
        {
            MacroFSNode item = macroTreeView.SelectedItem as MacroFSNode;
            string path = item.FullPath;

            FileSystemInfo file;
            string fileName = Path.GetFileNameWithoutExtension(path);
            string message;

            if (item.IsDirectory)
            {
                file = new DirectoryInfo(path);
                message = String.Format(VSMacros.Resources.DeleteFolder, fileName); 
            }
            else
            {
                file = new FileInfo(path);
                message = String.Format(VSMacros.Resources.DeleteMacro, fileName);
            }

            if (file.Exists)
            {
                if (MessageBoxResult.OK == MessageBox.Show(message, "Delete", MessageBoxButton.OKCancel, MessageBoxImage.Warning))
                {
                    file.Delete();  // TODO non-empty dir will raise an exception here -> User Directory.Delete(path, true)
                    item.Delete();
                }
            }
            else
            {
                item.Delete();
            }
        }

        private void SaveCurrentMacro(object sender, RoutedEventArgs e) { }
        private void Rename(object sender, RoutedEventArgs e)
        {
            MacroFSNode item = macroTreeView.SelectedItem as MacroFSNode;
            item.EnableEdit();
        }
        private void AssignShortcut(object sender, RoutedEventArgs e)
        {
        }

        private void Open(object sender, RoutedEventArgs e)
        {
            VSMacrosPackage.Current.OpenDirectory(null, null);
        }

        private void Refresh(object sender, RoutedEventArgs e)
        {
            MacroFSNode item = macroTreeView.SelectedItem as MacroFSNode;
            item.Refresh();
        }
        #endregion

        #region Events

        private void macroTreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            MacroFSNode SelectedNode = macroTreeView.SelectedItem as MacroFSNode;

            if (SelectedNode == null) return;

            SelectedPath = SelectedNode.FullPath;

            string name = Path.GetFileNameWithoutExtension(SelectedPath);
            string extension = Path.GetExtension(SelectedPath);

            if (extension != "")
            {
                if (name == "Current")
                {
                    macroTreeView.ContextMenu = macroTreeView.Resources["CurrentContext"] as System.Windows.Controls.ContextMenu;
                }
                else
                {
                    macroTreeView.ContextMenu = macroTreeView.Resources["MacroContext"] as System.Windows.Controls.ContextMenu;
                }
            }
            else if (name == "Macros")  // TODO change checking for something more robust
                macroTreeView.ContextMenu = macroTreeView.Resources["BrowserContext"] as System.Windows.Controls.ContextMenu;
            else
                macroTreeView.ContextMenu = macroTreeView.Resources["FolderContext"] as System.Windows.Controls.ContextMenu;

        }

        // Taken from http://stackoverflow.com/questions/592373/select-treeview-node-on-right-click-before-displaying-contextmenu/592483#592483
        private void OnTreeItemMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            TreeViewItem treeViewItem = VisualUpwardSearch(e.OriginalSource as DependencyObject);

            if (treeViewItem != null)
            {
                treeViewItem.Focus();
                e.Handled = true;
            }
        }

        static TreeViewItem VisualUpwardSearch(DependencyObject source)
        {
            while (source != null && !(source is TreeViewItem))
                source = VisualTreeHelper.GetParent(source);

            return source as TreeViewItem;
        }

        private void OnTreeKeyDown(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Delete:
                    {
                        Delete(null, null);
                        e.Handled = true;
                        break;
                    }
                case Key.Space:
                case Key.F2:
                    {
                        Rename(null, null);
                        e.Handled = true;
                        break;
                    }
            }
        }

        private void OnTextBoxLostFocus(object sender, RoutedEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            textBox.IsReadOnly = true;
        }

        #endregion
    }
}