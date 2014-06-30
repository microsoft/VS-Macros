using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design; // for CommandID
using System.IO;
using System.Runtime.InteropServices;
using VSMacros.Engines;
using VSMacros.Models;

namespace VSMacros
{
    [Guid(GuidList.GuidToolWindowPersistanceString)]
    public class MacrosToolWindow : ToolWindowPane
    {
        private const string CurrentMacroLocation = "Current.js";

        public MacrosToolWindow() :
            base(null)
        {
            this.Caption = Resources.ToolWindowTitle;
            this.BitmapResourceID = 301;
            this.BitmapIndex = 1;

            // Instantiate Tool Window Toolbar
            this.ToolBar = new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.MacrosToolWindowToolbar);

            // Initialize file system
            if (!Directory.Exists(VSMacrosPackage.Current.MacroDirectory))
            {
                Directory.CreateDirectory(VSMacrosPackage.Current.MacroDirectory);
            }

            if (!File.Exists(Path.Combine(VSMacrosPackage.Current.MacroDirectory, CurrentMacroLocation)))
            {
                File.Create(Path.Combine(VSMacrosPackage.Current.MacroDirectory, CurrentMacroLocation));
            }

            string macroDirectory = VSMacrosPackage.Current.MacroDirectory;

            // Create tree view root
            MacroFSNode root = new MacroFSNode(macroDirectory);

            // Make sure it is opened and selected by default
            root.IsExpanded = true;
            root.IsSelected = true;

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

        #endregion
    }
}
