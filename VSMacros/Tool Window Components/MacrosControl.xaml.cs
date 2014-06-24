using Microsoft.VisualStudio.Shell.Interop;
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
            InitializeComponent();
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
            MacroFSNode selectedNode = this.MacroTreeView.SelectedItem as MacroFSNode;

            if (selectedNode == null) 
            { 
                return; 
            }

            // TODO this is the way I have found to make the textBox disappear when it is shown
            // It might be a little costly
            // Instead, I should give the textBox focus when the template changers and monitor the LostFocus event
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
        #endregion
    }
}