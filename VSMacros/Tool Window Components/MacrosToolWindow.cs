//-----------------------------------------------------------------------
// <copyright file="MacrosToolWindow.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.ComponentModel.Design; // for CommandID
using System.IO;
using System.Runtime.InteropServices;
using EnvDTE80;
using Microsoft.Internal.VisualStudio.PlatformUI;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.PlatformUI;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using VSMacros.Engines;
using VSMacros.Interfaces;
using VSMacros.Models;

namespace VSMacros
{
    [Guid(GuidList.GuidToolWindowPersistanceString)]
    public class MacrosToolWindow : ToolWindowPane
    {
        private VSMacrosPackage owningPackage;
        private bool addedToolbarButton;

        public MacrosToolWindow() :
            base(null)
        {
            this.owningPackage = VSMacrosPackage.Current;
            this.Caption = Resources.ToolWindowTitle;
            this.BitmapResourceID = 301;
            this.BitmapIndex = 1;

            // Instantiate Tool Window Toolbar
            this.ToolBar = new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.MacrosToolWindowToolbar);

            Manager.CreateFileSystem();

            string macroDirectory = Manager.MacrosPath;

            // Create tree view root
            MacroFSNode root = new MacroFSNode(macroDirectory);

            // Make sure it is opened and selected by default
            root.IsExpanded = true;

            // Initialize Macros Control
            var macroControl = new MacrosControl(root);
            macroControl.Loaded += this.OnLoaded;
            this.Content = macroControl;

            Manager.Instance.LoadFolderExpansion();
        }

        public void OnLoaded(object sender, System.Windows.RoutedEventArgs e)
        {
            if (!this.addedToolbarButton)
            {
                IVsWindowFrame windowFrame = (IVsWindowFrame)GetService(typeof(SVsWindowFrame));

                object dteWindow;
                windowFrame.GetProperty((int)__VSFPROPID.VSFPROPID_ExtWindowObject, out dteWindow);
                Window2 window = (Window2)dteWindow;
                
                var controls = ((CommandBars)window.CommandBars)[1].Controls;

                this.owningPackage.ImageButtons.Add((CommandBarButton)controls[1]);
                this.owningPackage.ImageButtons.Add((CommandBarButton)controls[2]);
                this.owningPackage.ImageButtons.Add((CommandBarButton)controls[3]);

                this.addedToolbarButton = true;
            }

            IRecorderPrivate macroRecorder = (IRecorderPrivate)this.GetService(typeof(IRecorder));
            if (macroRecorder.IsRecording)
            {
                this.owningPackage.ChangeMenuIcons(this.owningPackage.StopIcon, 0);
            }
        }

        protected override void Initialize()
        {
            base.Initialize();

            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs)
            {
                // Create the command for refresh
                mcs.AddCommand(new MenuCommand(
                    this.Refresh,
                    new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdRefresh)));

                // Create the command for Collapse All
                mcs.AddCommand(new MenuCommand(
                    this.CollapseAll,
                    new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdCollapseAll)));

                // Create the command to open the macro directory
                mcs.AddCommand(new MenuCommand(
                    this.OpenDirectory,
                    new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdOpenDirectory)));

                // Create the command to open the selected folder
                mcs.AddCommand(new MenuCommand(
                    this.OpenSelectedFolder,
                    new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdOpenFolder)));

                // Create the command to edit a macro
                mcs.AddCommand(new MenuCommand(
                    this.Edit,
                    new CommandID(typeof(VSConstants.VSStd97CmdID).GUID, (int)VSConstants.VSStd97CmdID.Open)));

                // Create the command to assign a shortcut to a macro
                mcs.AddCommand(new MenuCommand(
                    this.AssignShortcut,
                    new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdAssignShortcut)));

                // Create the command to delete a macro
                mcs.AddCommand(new MenuCommand(
                    this.Delete,
                    new CommandID(typeof(VSConstants.VSStd97CmdID).GUID, (int)VSConstants.VSStd97CmdID.Delete)));

                // Create the command to rename a macro
                mcs.AddCommand(new MenuCommand(
                    this.Rename,
                    new CommandID(typeof(VSConstants.VSStd97CmdID).GUID, (int)VSConstants.VSStd97CmdID.Rename)));

                // Will bind F2 to the Rename handler
                mcs.AddCommand(new MenuCommand(
                this.Rename,
                new CommandID(typeof(VSConstants.VSStd97CmdID).GUID, (int)VSConstants.VSStd97CmdID.EditLabel)));

                // Create the command to rename a macro
                mcs.AddCommand(new MenuCommand(
                    this.NewMacro,
                    new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdNewMacro)));

                // Create the command to rename a macro
                mcs.AddCommand(new MenuCommand(
                    this.NewFolder,
                    new CommandID(typeof(VSConstants.VSStd97CmdID).GUID, (int)VSConstants.VSStd97CmdID.NewFolder)));
            }
        }

        /////////////////////////////////////////////////////////////////////////////
        // Command Handlers
        #region Command Handlers

        private void Refresh(object sender, EventArgs arguments)
        {
            Manager.Instance.Refresh();
        }

        private void CollapseAll(object sender, EventArgs arguments)
        {
            MacroFSNode.CollapseAllNodes(MacroFSNode.RootNode);
        }

        public void OpenDirectory(object sender, EventArgs arguments)
        {
            Manager.Instance.OpenFolder(Manager.MacrosPath);
        }

        public void OpenSelectedFolder(object sender, EventArgs arguments)
        {
            Manager.Instance.OpenFolder();
        }

        public void Edit(object sender, EventArgs arguments)
        {
            Manager.Instance.Edit();
        }

        public void AssignShortcut(object sender, EventArgs arguments)
        {
            Manager.Instance.AssignShortcut();
        }

        public void Delete(object sender, EventArgs arguments)
        {
            Manager.Instance.Delete();
        }

        public void Rename(object sender, EventArgs arguments)
        {
            Manager.Instance.Rename();
        }

        public void NewMacro(object sender, EventArgs arguments)
        {
            Manager.Instance.NewMacro();
        }

        public void NewFolder(object sender, EventArgs arguments)
        {
            Manager.Instance.NewFolder();
        }

        #endregion

        #region Search

        public override bool SearchEnabled
        {
            get { return true; }
        }

        public override IVsSearchTask CreateSearch(uint cookie, IVsSearchQuery searchQuery, IVsSearchCallback searchCallback)
        {
            if (searchQuery == null || searchCallback == null)
                return null;
            return new SearchTask(cookie, searchQuery, searchCallback, this);
        }

        public override void ClearSearch()
        {
            MacroFSNode.DisableSearch();
        }

        public override void ProvideSearchSettings(IVsUIDataSource searchSettings)
        {
            Utilities.SetValue(searchSettings, SearchSettingsDataSource.SearchStartTypeProperty.Name, (uint)VSSEARCHSTARTTYPE.SST_INSTANT);
            Utilities.SetValue(searchSettings, SearchSettingsDataSource.SearchProgressTypeProperty.Name, (uint)VSSEARCHPROGRESSTYPE.SPT_DETERMINATE);
        }

        private IVsEnumWindowSearchOptions searchOptionsEnum;
        public override IVsEnumWindowSearchOptions SearchOptionsEnum
        {
            get
            {
                if (this.searchOptionsEnum == null)
                {
                    List<IVsWindowSearchOption> list = new List<IVsWindowSearchOption>();

                    list.Add(this.MatchCaseOption);
                    list.Add(this.WithinFileOption);

                    this.searchOptionsEnum = new WindowSearchOptionEnumerator(list) as IVsEnumWindowSearchOptions;
                }
                return this.searchOptionsEnum;
            }
        }

        private WindowSearchBooleanOption withinFileOption;
        public WindowSearchBooleanOption WithinFileOption
        {
            get
            {
                if (this.withinFileOption == null)
                {
                    this.withinFileOption = new WindowSearchBooleanOption(Resources.SearchWithinFileContents, Resources.SearchWithinFileContents, false);
                }
                return this.withinFileOption;
            }
        }

        private WindowSearchBooleanOption matchCaseOption;
        public WindowSearchBooleanOption MatchCaseOption
        {
            get
            {
                if (this.matchCaseOption == null)
                {
                    this.matchCaseOption = new WindowSearchBooleanOption(Resources.MatchCase, Resources.MatchCase, false);
                }
                return this.matchCaseOption;
            }
        }

        internal class SearchTask : VsSearchTask
        {
            private MacrosToolWindow toolWindow;

            public SearchTask(uint cookie, IVsSearchQuery searchQuery, IVsSearchCallback searchCallback, MacrosToolWindow toolwindow)
                : base(cookie, searchQuery, searchCallback)
            {
                this.toolWindow = toolwindow;
            }

            protected override void OnStartSearch()
            {
                // Enable search on the MacroFSNodes so that only the matched nodes will be shown
                MacroFSNode.EnableSearch();

                // Get the search option. 
                bool matchCase = this.toolWindow.MatchCaseOption.Value;
                bool withinFileContents = this.toolWindow.WithinFileOption.Value;

                try
                {
                    string searchString = this.SearchQuery.SearchString;
                    StringComparison comp = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

                    this.TraverseAndMark(MacroFSNode.RootNode, searchString, comp, withinFileContents);
                }
                catch (Exception e)
                {
                    System.Windows.Forms.MessageBox.Show(e.Message);
                    this.ErrorCode = VSConstants.E_FAIL;
                }

                // Call the implementation of this method in the base class. 
                // This sets the task status to complete and reports task completion. 
                base.OnStartSearch();
            }

            protected override void OnStopSearch()
            {
                // Disable the search (will notify all nodes of the change)
                MacroFSNode.DisableSearch();
            }

            private void TraverseAndMark(MacroFSNode root, string searchString, StringComparison comp, bool withinFileContents)
            {
                if (this.Contains(root.Name, searchString, comp))
                {
                    root.IsMatch = true;
                }
                else if (withinFileContents && !root.IsDirectory)
                {
                    System.Threading.Tasks.Task.Run(() =>
                    {
                        string allText = File.ReadAllText(root.FullPath);
                        if (this.Contains(allText, searchString, comp))
                        {
                            root.IsMatch = true;
                        }
                    });
                }
                else
                {
                    root.IsMatch = false;
                }

                if (root.Children != null)
                {
                    foreach (var child in root.Children)
                    {
                        this.TraverseAndMark(child, searchString, comp, withinFileContents);
                    }
                }
            }

            private bool Contains(string source, string toCheck, StringComparison comp)
            {
                return source.IndexOf(toCheck, comp) >= 0;
            }
        }

        #endregion
    }
}
