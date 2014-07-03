using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design; // for CommandID
using System.IO;
using System.Runtime.InteropServices;
using VSMacros.Engines;
using VSMacros.Models;
using Microsoft.VisualStudio.Shell.Interop;
using System.Windows.Controls;
using Microsoft.VisualStudio.PlatformUI;
using Microsoft.Internal.VisualStudio.PlatformUI;
using System.Collections.Generic;

namespace VSMacros
{
    [Guid(GuidList.GuidToolWindowPersistanceString)]
    public class MacrosToolWindow : ToolWindowPane
    {
        public MacrosToolWindow() :
            base(null)
        {
            this.Caption = Resources.ToolWindowTitle;
            this.BitmapResourceID = 301;
            this.BitmapIndex = 1;

            // Instantiate Tool Window Toolbar
            this.ToolBar = new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.MacrosToolWindowToolbar);

            Manager.Instance.CreateFileSystem();

            string macroDirectory = VSMacrosPackage.Current.MacroDirectory;

            // Create tree view root
            MacroFSNode root = new MacroFSNode(macroDirectory);

            // Make sure it is opened and selected by default
            root.IsExpanded = true;

            // Initialize Macros Control
            base.Content = new MacrosControl(root);
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
                    new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdEdit)));

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

        public string MacroDirectory
        {
            get
            {
                return VSMacrosPackage.Current.MacroDirectory;
            }
        }

        /////////////////////////////////////////////////////////////////////////////
        // Command Handlers
        #region Command Handlers

        private void Refresh(object sender, EventArgs arguments)
        {
            Manager.Instance.Refresh();
        }

        public void OpenDirectory(object sender, EventArgs arguments)
        {
            Manager.Instance.OpenFolder(this.MacroDirectory);
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

        public override IVsSearchTask CreateSearch(uint dwCookie, IVsSearchQuery pSearchQuery, IVsSearchCallback pSearchCallback)
        {
            if (pSearchQuery == null || pSearchCallback == null)
                return null;
            return new SearchTask(dwCookie, pSearchQuery, pSearchCallback, this);
        }

        public override void ClearSearch()
        {
            MacroFSNode.DisableSearch();
        }

        public override void ProvideSearchSettings(IVsUIDataSource pSearchSettings)
        {
            Utilities.SetValue(pSearchSettings, SearchSettingsDataSource.SearchStartTypeProperty.Name, (uint)VSSEARCHSTARTTYPE.SST_INSTANT);
            Utilities.SetValue(pSearchSettings, SearchSettingsDataSource.SearchProgressTypeProperty.Name, (uint)VSSEARCHPROGRESSTYPE.SPT_DETERMINATE);
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
                    this.withinFileOption = new WindowSearchBooleanOption("Search within file contents", "Search within file contents", false);
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
                    this.matchCaseOption = new WindowSearchBooleanOption("Match case", "Match case", false);
                }
                return this.matchCaseOption;
            }
        }

        internal class SearchTask : VsSearchTask
        {
            private MacrosToolWindow toolWindow;

            public SearchTask(uint dwCookie, IVsSearchQuery pSearchQuery, IVsSearchCallback pSearchCallback, MacrosToolWindow toolwindow)
                : base(dwCookie, pSearchQuery, pSearchCallback)
            {
                this.toolWindow = toolwindow;
            }

            protected override void OnStartSearch()
            {
                MacroFSNode.EnableSearch();

                // Get the search option. 
                bool matchCase = toolWindow.MatchCaseOption.Value;
                bool withinFileContents = toolWindow.WithinFileOption.Value;

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
                MacroFSNode.DisableSearch();
            }


            private void TraverseAndMark(MacroFSNode root, string searchString, StringComparison comp, bool withinFileContents)
            {
                if (this.Contains(root.FullPath, searchString, comp))
                {
                    root.IsMatch = true;
                }
                else if (withinFileContents && !root.IsDirectory)
                {
                    // TODO move to b/g thread!
                    string allText = File.ReadAllText(root.FullPath);
                    if (this.Contains(allText, searchString, comp))
                    {
                        root.IsMatch = true;
                    }
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
