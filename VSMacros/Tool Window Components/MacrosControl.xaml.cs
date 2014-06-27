// MacrosControl.xaml.cs

using System;
using System.Collections.Generic;
using System.IO;
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
using Microsoft.VisualStudio.Shell.Interop;
using VSMacros.Models;

namespace VSMacros
{
    public partial class MacrosControl : UserControl
    {
        public static MacrosControl Current { get; private set; }

        public MacrosControl(MacroFSNode rootNode)
        {
            Current = this;

            MacroFSNode.RootNode = rootNode;

            // Let the UI bind to the view-model
            this.DataContext = new MacroFSNode[] { rootNode };
            this.InitializeComponent();
        }

        public MacroFSNode SelectedNode
        {
            get
            {
                return this.MacroTreeView.SelectedItem as MacroFSNode;
            }
        }

        private TreeViewItem GetSelectedTreeViewItem()
        {
            return (TreeViewItem)(this.MacroTreeView.ItemContainerGenerator.ContainerFromIndex(this.MacroTreeView.Items.CurrentPosition));
        }

        #region Events

        private void MacroTreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            // TODO this is the way I have found to make the textBox disappear when it is shown
            // It might be a little expensive
            // Instead, I should give the textBox focus when the template changes and monitor the LostFocus event
            MacroFSNode oldNode = e.OldValue as MacroFSNode;
            if (oldNode != null)
            {
                oldNode.IsEditable = false;
            }
        }

        private void TreeViewItem_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            // Make sure that the clicks has selected the item
            TreeViewItem treeViewItem = VisualUpwardSearch(e.OriginalSource as DependencyObject);

            if (treeViewItem != null)
            {
                treeViewItem.Focus();
                e.Handled = true;
            }

            /* Show Context Menu */
            IVsUIShell uiShell = (IVsUIShell)((IServiceProvider)VSMacrosPackage.Current).GetService(typeof(SVsUIShell));

            if (uiShell != null)
            {
                // Get CmdID
                MacroFSNode selectedNode = this.MacroTreeView.SelectedItem as MacroFSNode;
                int menuID;

                if (selectedNode.IsDirectory)
                {
                    if (selectedNode == MacroFSNode.RootNode)
                    {
                        menuID = PkgCmdIDList.BrowserContextMenu;
                    }
                    else
                    {
                        menuID = PkgCmdIDList.FolderContextMenu;
                    }
                }
                else
                {
                    if (selectedNode.Name == "Current")
                    {
                        menuID = PkgCmdIDList.CurrentContextMenu;
                    }
                    else
                    {
                        menuID = PkgCmdIDList.MacroContextMenu;
                    }
                }

                System.Drawing.Point pt = System.Windows.Forms.Cursor.Position;
                POINTS[] pnts = new POINTS[1];
                pnts[0].x = (short)pt.X;
                pnts[0].y = (short)pt.Y;

                uiShell.ShowContextMenu(0, GuidList.GuidVSMacrosCmdSet, menuID, pnts, null);
            }
        }

        private static TreeViewItem VisualUpwardSearch(DependencyObject source)
        {
            while (source != null && !(source is TreeViewItem))
            {
                source = VisualTreeHelper.GetParent(source);
            }

            return source as TreeViewItem;
        }

        private void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                this.SelectedNode.DisableEdit();
            }
        }

        private void TextBox_LostKeyboardFocus(object sender, RoutedEventArgs e)
        {
            this.SelectedNode.DisableEdit();
        }

        #endregion

        #region Drag & Drop
        private Point startPos;

        private void TreeViewItem_PreviewMouseLeftButtonDown(object sender, MouseEventArgs e)
        {
            // Save current position so we have a reference to compare against in TreeViewItem_MouseMove
            this.startPos = e.GetPosition(null);
        }
        private void TreeViewItem_MouseMove(object sender, MouseEventArgs e)
        {
            return;
            System.Diagnostics.Debug.WriteLine("In mouve move");
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                // Has the mouse move enough?
                var mousePos = e.GetPosition(null);
                var diff = mousePos - this.startPos;

                if (Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance ||
                    Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance)
                {
                    // Get the dragged TreeViewItem
                    TreeView tv = sender as TreeView;
                    TreeViewItem tvi = FindAnchestor<TreeViewItem>((DependencyObject)e.OriginalSource);

                    // Find the data behind the TreeViewItem
                    MacroFSNode node = tvi.Header as MacroFSNode;

                    // Initialize the drag & drop operation
                    MacroFSNode data = this.MacroTreeView.SelectedItem as MacroFSNode;
                    DragDrop.DoDragDrop(this.MacroTreeView, data, DragDropEffects.Move);
                }
            }
        }

        // Helper to search up the VisualTree
        private static T FindAnchestor<T>(DependencyObject current)
            where T : DependencyObject
        {
            do
            {
                if (current is T)
                {
                    return (T)current;
                }
                current = VisualTreeHelper.GetParent(current);
            }
            while (current != null);
            return null;
        }

        private void TreeViewItem_Drop(object sender, DragEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("In TVI Drop");
            if (e.Data.GetDataPresent(typeof(MacroFSNode)))
            {
                MacroFSNode node = e.Data.GetData(typeof(MacroFSNode)) as MacroFSNode;
            }
        }

        private void TreeViewItem_DragEnter(object sender, DragEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("In TVI Drag enter");
            if (!e.Data.GetDataPresent(typeof(MacroFSNode)))
            {
                e.Effects = DragDropEffects.None;
            }
        }

        #endregion

        private void MacroTreeView_DragOver(object sender, DragEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("In TVI over");
            if ((e.AllowedEffects & DragDropEffects.Move) == DragDropEffects.Move)
            {
                e.Effects = DragDropEffects.Move;
            }
        }
    }
}