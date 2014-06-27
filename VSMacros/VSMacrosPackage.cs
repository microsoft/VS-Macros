﻿using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using VSMacros.Engines;

namespace VSMacros
{
    
    [ProvideToolWindow(typeof(MacrosToolWindow), Style = VsDockStyle.Tabbed, Window = "3ae79031-e1bc-11d0-8f78-00a0c9110057")]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid(GuidList.GuidVSMacrosPkgString)]
    public sealed class VSMacrosPackage : Package
    {
        public static VSMacrosPackage Current { get; private set; }

        public VSMacrosPackage()
        {
            Current = this;
        }

        private string macroDirectory;
        public string MacroDirectory
        {
            get 
            { 
                if (this.macroDirectory == default(string))
                    this.macroDirectory = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Macros");
                return this.macroDirectory;
            }
        }

        /////////////////////////////////////////////////////////////////////////////
        // Overridden Package Implementation
        #region Package Members

        protected override void Initialize()
        {
            base.Initialize();

            // QUESTION Should some of the commands, namely those only appearing in the tool window, be added later, maybe when the tool window is created?
            // Add our command handlers for the menu
            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs)
            {
                // Create the command for the tool window
                mcs.AddCommand(new MenuCommand(
                   this.ShowToolWindow,
                   new CommandID(GuidList.GuidVSMacrosCmdSet, (int)PkgCmdIDList.CmdIdMacroExplorer)));

                 // Create the command for start recording
                mcs.AddCommand(new MenuCommand(
                  this.Record,
                  new CommandID(GuidList.GuidVSMacrosCmdSet, (int)PkgCmdIDList.CmdIdRecord)));

                // Create the command for playback
                mcs.AddCommand(new MenuCommand(
                    this.Playback,
                    new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdPlayback)));

                // Create the command for playback multiple times
                mcs.AddCommand(new MenuCommand(
                    this.PlaybackMultipleTimes,
                    new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdPlaybackMultipleTimes)));

                // Create the command for save current macro
                mcs.AddCommand(new MenuCommand(
                    this.SaveCurrent,
                    new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdSaveTemporaryMacro)));

                // Create the command for refresh
                mcs.AddCommand(new MenuCommand(
                    this.Refresh,
                    new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdRefresh)));

                // Create the command to open the macro directory
                mcs.AddCommand(new MenuCommand(
                    this.OpenDirectory,
                    new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdOpenDirectory)));

                // Create the command to edit a macro
                mcs.AddCommand(new MenuCommand(
                    this.Edit,
                    new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdEdit)));

                // Create the command to rename a macro
                mcs.AddCommand(new MenuCommand(
                    this.Rename,
                    new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdRename)));

                // Create the command to assign a shortcut to a macro
                mcs.AddCommand(new MenuCommand(
                    this.AssignShortcut,
                    new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdAssignShortcut)));

                // Create the command to delete a macro
                mcs.AddCommand(new MenuCommand(
                    this.Delete,
                    new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdDelete)));

                // Create the command to playback bounded macros
                mcs.AddCommand(new MenuCommand(this.PlaybackCommand1, new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdCommand1)));
                mcs.AddCommand(new MenuCommand(this.PlaybackCommand2, new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdCommand2)));
                mcs.AddCommand(new MenuCommand(this.PlaybackCommand3, new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdCommand3)));
                mcs.AddCommand(new MenuCommand(this.PlaybackCommand4, new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdCommand4)));
                mcs.AddCommand(new MenuCommand(this.PlaybackCommand5, new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdCommand5)));
                mcs.AddCommand(new MenuCommand(this.PlaybackCommand6, new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdCommand6)));
                mcs.AddCommand(new MenuCommand(this.PlaybackCommand7, new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdCommand7)));
                mcs.AddCommand(new MenuCommand(this.PlaybackCommand8, new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdCommand8)));
                mcs.AddCommand(new MenuCommand(this.PlaybackCommand9, new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdCommand9)));
            }
        }
        #endregion

        /////////////////////////////////////////////////////////////////////////////
        // Command Handlers
        #region Command Handlers

        private void ShowToolWindow(object sender = null, EventArgs e = null)
        {
            // Get the (only) instance of this tool window
            // The last flag is set to true so that if the tool window does not exists it will be created.
            ToolWindowPane window = this.FindToolWindow(typeof(MacrosToolWindow), 0, true);
            if ((window == null) || (window.Frame == null))
            {
                throw new NotSupportedException(Resources.CannotCreateWindow);
            }
            IVsWindowFrame windowFrame = (IVsWindowFrame)window.Frame;
            Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(windowFrame.Show());
        }

        public void Record(object sender, EventArgs arguments)
        {
            Manager.Instance.ToggleRecording();
        }

        private void Playback(object sender, EventArgs arguments)
        {
            Manager.Instance.Playback(string.Empty, 1);
        }

        private void PlaybackMultipleTimes(object sender, EventArgs arguments)
        {
            Manager.Instance.Playback(string.Empty, 0);
        }

        private void SaveCurrent(object sender, EventArgs arguments)
        {
            Manager.Instance.SaveCurrent();
        }

        private void Refresh(object sender, EventArgs arguments)
        {
            Manager.Instance.Refresh();
        }

        public void OpenDirectory(object sender, EventArgs arguments)
        {
            // Open the macro directory and let the user manage the macros
            System.Threading.Tasks.Task.Run(() => { System.Diagnostics.Process.Start(MacroDirectory); });
        }

        public void Edit(object sender, EventArgs arguments)
        {
            Manager.Instance.Edit();
        }

        public void Rename(object sender, EventArgs arguments)
        {
            Manager.Instance.Rename();
        }

        public void AssignShortcut(object sender, EventArgs arguments)
        {
            Manager.Instance.AssignShortcut();
        }

        public void Delete(object sender, EventArgs arguments)
        {
            Manager.Instance.Delete();
        }

        public void PlaybackCommand1(object sender, EventArgs arguments) { Manager.Instance.PlaybackCommand("command1");}
        public void PlaybackCommand2(object sender, EventArgs arguments) { Manager.Instance.PlaybackCommand("command2"); }
        public void PlaybackCommand3(object sender, EventArgs arguments) { Manager.Instance.PlaybackCommand("command3"); }
        public void PlaybackCommand4(object sender, EventArgs arguments) { Manager.Instance.PlaybackCommand("command4"); }
        public void PlaybackCommand5(object sender, EventArgs arguments) { Manager.Instance.PlaybackCommand("command5"); }
        public void PlaybackCommand6(object sender, EventArgs arguments) { Manager.Instance.PlaybackCommand("command6"); }
        public void PlaybackCommand7(object sender, EventArgs arguments) { Manager.Instance.PlaybackCommand("command7"); }
        public void PlaybackCommand8(object sender, EventArgs arguments) { Manager.Instance.PlaybackCommand("command8"); }
        public void PlaybackCommand9(object sender, EventArgs arguments) { Manager.Instance.PlaybackCommand("command9"); }

        #endregion
    }
}
