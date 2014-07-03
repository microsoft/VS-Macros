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

        private TreeViewItem GetSelectedTreeViewItem()
        {
            return (TreeViewItem)(this.MacroTreeView.ItemContainerGenerator.ContainerFromIndex(this.MacroTreeView.Items.CurrentPosition));
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
            MacroFSNode.FindNodeFromFullPath(Path.Combine(VSMacrosPackage.Current.MacroDirectory, "Current.js")).IsSelected = true;
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
                TextBox textBox = sender as TextBox;

                BindingOperations.GetBindingExpression(textBox, TextBox.TextProperty).UpdateSource();
                this.SelectedNode.DisableEdit();
            }
        }

        private void TextBox_LostKeyboardFocus(object sender, RoutedEventArgs e)
        {
            Keyboard.ClearFocus();
            this.SelectedNode.DisableEdit();
        }

        private void TreeViewItem_MouseDoubleClick(object sender, RoutedEventArgs e)
        {
            if (!this.SelectedNode.IsDirectory)
            {
                VSMacros.Engines.Manager.Instance.Playback(this.SelectedNode.FullPath, 1);
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
        private MacroFSNode targetNode;
        private MacroFSNode draggedNode;
        public static readonly DependencyProperty IsTreeViewItemDropOverProperty = DependencyProperty.RegisterAttached("IsTreeViewItemDropOver", typeof(bool), typeof(MacrosControl), new PropertyMetadata(false));

        private void TreeViewItem_PreviewMouseLeftButtonDown(object sender, MouseEventArgs e)
        {
            // Save current position so we have a reference to compare against in TreeViewItem_MouseMove
            this.startPos = e.GetPosition(this.MacroTreeView);
        }

        private void TreeViewItem_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                // Has the mouse move enough?
                var mousePos = e.GetPosition(this.MacroTreeView);
                var diff = mousePos - this.startPos;

                if (Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance ||
                    Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance)
                {
                    // Get the dragged TreeViewItem
                    this.draggedNode = this.MacroTreeView.SelectedItem as MacroFSNode;

                    // The root node is not dragable
                    if (this.draggedNode == MacroFSNode.RootNode)
                    {
                        this.draggedNode = null;
                    }

                    if (this.draggedNode != null)
                    {
                        // Initialize the drag & drop operation
                        DragDropEffects finalDropEffect = DragDrop.DoDragDrop(this.MacroTreeView, this.draggedNode, DragDropEffects.Move);

                        // Checking target is not null and item is dragging
                        if ((finalDropEffect == DragDropEffects.Move) && (this.targetNode != null))
                        {
                            // A Move drop is accepted
                            if (!this.draggedNode.Equals(this.targetNode))
                            {
                                this.MoveItem(this.draggedNode, this.targetNode);
                                this.targetNode = null;
                                this.draggedNode = null;
                            }
                        }
                    }                   
                }
            }
        }

        private void MacroTreeView_DragOver(object sender, DragEventArgs e)
        {
            Point currentPosition = e.GetPosition(this.MacroTreeView);

            if ((Math.Abs(currentPosition.X - this.startPos.X) > SystemParameters.MinimumHorizontalDragDistance) ||
                (Math.Abs(currentPosition.Y - this.startPos.Y) > SystemParameters.MinimumVerticalDragDistance))
            {
                // Verify that this is a valid drop
                TreeViewItem item = this.GetNearestContainer(e.OriginalSource as UIElement);
                if (this.ValidDropTarget(this.draggedNode, item.Header as MacroFSNode))
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

            // Verify that this is a valid drop and then store the drop target
            TreeViewItem targetItem = this.GetNearestContainer(e.OriginalSource as UIElement);
            if (targetItem != null && this.draggedNode != null)
            {
                MacroFSNode targetNode = targetItem.Header as MacroFSNode;

                if (!targetNode.IsDirectory)
                {
                    this.targetNode = targetNode.Parent;
                }
                else
                {
                    this.targetNode = targetNode;
                }

                e.Effects = DragDropEffects.Move;
            }
        }

        private void TreeViewItem_DragEnter(object sender, DragEventArgs e)
        {
            // Highlight item on DragEnter
            TreeViewItem item = sender as TreeViewItem;
            SetIsTreeViewItemDropOver(item, true);
            e.Handled = true;
        }

        private void TreeViewItem_DragLeave(object sender, DragEventArgs e)
        {
            // Remove highlight on DragLeave
            TreeViewItem item = sender as TreeViewItem;
            SetIsTreeViewItemDropOver(item, false);
            e.Handled = true;
        }

        public static bool GetIsTreeViewItemDropOver(TreeViewItem item)
        {
            return (bool)item.GetValue(IsTreeViewItemDropOverProperty);
        }

        public static void SetIsTreeViewItemDropOver(TreeViewItem item, bool value)
        {
            item.SetValue(IsTreeViewItemDropOverProperty, value);
        }

        private void MoveItem(MacroFSNode sourceItem, MacroFSNode targetItem)
        {
            try
            {
                string sourcePath = sourceItem.FullPath;
                string targetPath = System.IO.Path.Combine(targetItem.FullPath, sourceItem.Name);
                string extension = ".js";

                // Move on disk
                if (sourceItem.IsDirectory)
                {
                    System.IO.Directory.Move(sourcePath, targetPath);
                }
                else
                {
                    targetPath = targetPath + extension;
                    System.IO.File.Move(sourcePath, targetPath);
                }

                // Move shortcut as well
                if (!string.IsNullOrEmpty(sourceItem.Shortcut))
                {
                    int shortcutNumber = sourceItem.Shortcut[sourceItem.Shortcut.Length - 2] - '0';
                    Manager.Instance.Shortcuts[shortcutNumber] = targetPath;
                }

                // Refresh tree
                MacroFSNode.RefreshTree();

                // Restore previously selected node
                MacroFSNode selected = MacroFSNode.FindNodeFromFullPath(targetPath);
                selected.IsSelected = true;
                
                // and notify change in shortcut
                selected.Shortcut = null;
            }
            catch (Exception e)
            {
                Manager.Instance.ShowMessageBox(e.Message);
            }
        }

        private bool ValidDropTarget(MacroFSNode sourceItem, MacroFSNode targetItem)
        {
            // Check whether the target item is meeting your condition
            if (!sourceItem.Equals(targetItem))
            {
                return true;
            }
            return false;
        }

        private static TObject FindVisualParent<TObject>(UIElement child) where TObject : UIElement
        {
            if (child == null)
            {
                return null;
            }

            UIElement parent = VisualTreeHelper.GetParent(child) as UIElement;

            while (parent != null)
            {
                TObject found = parent as TObject;
                if (found != null)
                {
                    return found;
                }
                else
                {
                    parent = VisualTreeHelper.GetParent(parent) as UIElement;
                }
            }

            return null;
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