//-----------------------------------------------------------------------
// <copyright file="MacroControl.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using VSMacros.Engines;
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

        #region Events

        private void MacroTreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            // Disable edit for previously selected node
            MacroFSNode oldNode = e.OldValue as MacroFSNode;
            if (oldNode != null)
            {
                oldNode.DisableEdit();
            }
        }

        private void MacroTreeView_Loaded(object sender, RoutedEventArgs e)
        {
            // Select Current macro
            MacroFSNode.FindNodeFromFullPath(Manager.CurrentMacroPath).IsSelected = true;
        }

        private void TreeViewItem_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            // Make sure dragging is not initiated
            this.isDragging = false;

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
                // Get context menu id
                int menuID;
                MacroFSNode selectedNode = this.MacroTreeView.SelectedItem as MacroFSNode;

                if (selectedNode.IsDirectory)
                {
                    if (selectedNode == MacroFSNode.RootNode)
                    {
                        menuID = PkgCmdIDList.BrowserContextMenu;
                    }
                    else if (this.InSamples(selectedNode))
                    {
                        menuID = PkgCmdIDList.SampleFolderContextMenu;
                    }
                    else
                    {
                        menuID = PkgCmdIDList.FolderContextMenu;
                    }
                }
                else
                {
                    if (selectedNode.FullPath == Manager.CurrentMacroPath)
                    {
                        menuID = PkgCmdIDList.CurrentContextMenu;
                    }
                    else if (this.InSamples(selectedNode))
                    {
                        menuID = PkgCmdIDList.SampleMacroContextMenu;
                    }
                    else
                    {
                        menuID = PkgCmdIDList.MacroContextMenu;
                    }
                }

                // Show right context menu
                System.Drawing.Point pt = System.Windows.Forms.Cursor.Position;
                POINTS[] pnts = new POINTS[1];
                pnts[0].x = (short)pt.X;
                pnts[0].y = (short)pt.Y;

                uiShell.ShowContextMenu(0, GuidList.GuidVSMacrosCmdSet, menuID, pnts, null);
            }
        }

        private bool InSamples(MacroFSNode node)
        {
            do
            {
                if (node.FullPath == Manager.SamplesFolderPath)
                {
                    return true;
                }

                node = node.Parent;
            } while (node != MacroFSNode.RootNode);

            return false;
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
            // Update source and exit edit on Enter
            if (e.Key == Key.Enter)
            {
                TextBox textBox = sender as TextBox;

                // Update source
                BindingOperations.GetBindingExpression(textBox, TextBox.TextProperty).UpdateSource();

                // Disable edit for selected macro
                this.SelectedNode.DisableEdit();
            }
        }

        private void TreeViewItem_MouseDoubleClick(object sender, RoutedEventArgs e)
        {
            MacroFSNode node = ((TreeViewItem)sender).Header as MacroFSNode;

            if (!node.IsDirectory)
            {
                Action playback = () => Manager.Instance.Playback(node.FullPath);
                this.Dispatcher.BeginInvoke(playback);
            }
        }

        private void TextBox_Loaded(object sender, RoutedEventArgs e)
        {
            TextBox textBox = sender as TextBox;

            textBox.SelectAll();
            textBox.Focus();
        }

        #endregion

        #region Drag & Drop
        private Point startPos;
        private bool isDragging;
        public static readonly DependencyProperty IsTreeViewItemDropOverProperty = DependencyProperty.RegisterAttached("IsTreeViewItemDropOver", typeof(bool), typeof(MacrosControl), new PropertyMetadata(false));

        private void TreeViewItem_PreviewMouseLeftButtonDown(object sender, MouseEventArgs e)
        {
            this.isDragging = true;

            // Save current position so we have a reference to compare against in TreeViewItem_MouseMove
            this.startPos = e.GetPosition(this.MacroTreeView);
        }

        private void TreeViewItem_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                // Has the mouse moved enough?
                var mousePos = e.GetPosition(this.MacroTreeView);
                var diff = mousePos - this.startPos;

                if (Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance ||
                    Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance &&
                    this.isDragging)
                {
                    // Get the dragged MacroFSNode
                    MacroFSNode draggedNode = this.MacroTreeView.SelectedItem as MacroFSNode;

                    // The root node is not draggable
                    if (draggedNode != null && draggedNode != MacroFSNode.RootNode)
                    {
                        // Initialize the drag & drop operation
                        DragDrop.DoDragDrop(this.MacroTreeView, draggedNode, DragDropEffects.Move);
                    }
                }
            }
        }

        private void MacroTreeView_DragOver(object sender, DragEventArgs e)
        {
            Point currentPosition = e.GetPosition(this.MacroTreeView);

            if ((Math.Abs(currentPosition.X - this.startPos.X) > SystemParameters.MinimumHorizontalDragDistance) ||
                (Math.Abs(currentPosition.Y - this.startPos.Y) > SystemParameters.MinimumVerticalDragDistance) &&
                this.isDragging &&
                e.Data.GetDataPresent(typeof(MacroFSNode)) || e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                // Verify that this is a valid drop
                TreeViewItem item = this.GetNearestContainer(e.OriginalSource as UIElement);
                MacroFSNode node = item.Header as MacroFSNode;
                if (node != null && !this.InSamples(node))
                {
                    e.Effects = DragDropEffects.Move;
                }
                else
                {
                    e.Effects = DragDropEffects.None;
                }
            }
            e.Handled = true;
        }

        private void TreeViewItem_Drop(object sender, DragEventArgs e)
        {
            e.Effects = DragDropEffects.None;
            e.Handled = true;

            // Get target node
            TreeViewItem targetItem = this.GetNearestContainer(e.OriginalSource as UIElement);
            MacroFSNode target = targetItem.Header as MacroFSNode;

            if (e.Data.GetDataPresent(typeof(MacroFSNode)) && this.isDragging)
            {
                // Get dragged node
                MacroFSNode dragged = e.Data.GetData(typeof(MacroFSNode)) as MacroFSNode;

                if (target != null && dragged != null && target != dragged)
                {
                    if (!target.IsDirectory)
                    {
                        target = target.Parent;
                    }

                    Manager.Instance.MoveItem(dragged, target);
                }
            }
            else if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] droppedFiles = e.Data.GetData(DataFormats.FileDrop, true) as string[];

                if (!target.IsDirectory)
                {
                    target = target.Parent;
                }

                foreach (string s in droppedFiles)
                {
                    string filename = Path.GetFileNameWithoutExtension(s);

                    // If the file is a macro file
                    if (Path.GetExtension(s) == ".js")
                    {
                        string destPath = Path.Combine(target.FullPath, filename + ".js");
                        VSConstants.MessageBoxResult result = VSConstants.MessageBoxResult.IDYES;

                        if (File.Exists(destPath))
                        {
                            string message = string.Format(VSMacros.Resources.DragDropFileExists, filename);
                            result = Manager.Instance.ShowMessageBox(message, OLEMSGBUTTON.OLEMSGBUTTON_YESNO);
                        }

                        if (result == VSConstants.MessageBoxResult.IDYES)
                        {
                            File.Copy(s, destPath, true);
                        }
                    }
                    else if (Path.GetExtension(s) == "")
                    {
                        string destPath = Path.Combine(target.FullPath, filename);
                        VSConstants.MessageBoxResult result = VSConstants.MessageBoxResult.IDYES;

                        if (Directory.Exists(destPath))
                        {
                            string message = string.Format(VSMacros.Resources.DragDropFileExists, filename);
                            result = Manager.Instance.ShowMessageBox(message, OLEMSGBUTTON.OLEMSGBUTTON_YESNO);
                        }

                        if (result == VSConstants.MessageBoxResult.IDYES)
                        {
                            Manager.DirectoryCopy(s, destPath, true);
                        }
                    }

                }

                // Unset IsTreeViewItemDropOver for target
                MacrosControl.SetIsTreeViewItemDropOver(targetItem, false);

                MacroFSNode.RefreshTree(target);
            }

            this.isDragging = false;
        }

        private void TreeViewItem_DragEnter(object sender, DragEventArgs e)
        {
            if (this.isDragging || e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                // Highlight item on DragEnter
                TreeViewItem item = sender as TreeViewItem;

                if ((MacroFSNode)item.Header != MacroFSNode.RootNode)
                {
                    MacrosControl.SetIsTreeViewItemDropOver(item, true);
                }
                e.Handled = true;
            }
        }

        private void TreeViewItem_DragLeave(object sender, DragEventArgs e)
        {
            if (this.isDragging || e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                // Remove highlight on DragLeave
                TreeViewItem item = sender as TreeViewItem;
                MacrosControl.SetIsTreeViewItemDropOver(item, false);
                e.Handled = true;
            }
        }

        public static bool GetIsTreeViewItemDropOver(TreeViewItem item)
        {
            return (bool)item.GetValue(IsTreeViewItemDropOverProperty);
        }

        public static void SetIsTreeViewItemDropOver(TreeViewItem item, bool value)
        {
            item.SetValue(IsTreeViewItemDropOverProperty, value);
        }

        private bool ValidDropTarget(MacroFSNode sourceItem, MacroFSNode targetItem)
        {
            // Check whether the target item is meeting your condition
            return !sourceItem.Equals(targetItem);
        }

        private TreeViewItem GetNearestContainer(UIElement element)
        {
            // Walk up the element tree to the nearest tree view item.
            TreeViewItem container = element as TreeViewItem;
            while ((container == null) && (element != null))
            {
                element = VisualTreeHelper.GetParent(element) as UIElement;
                container = element as TreeViewItem;
            }
            return container;
        }

        #endregion
    }
}