using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using VSMacros.Engines;

namespace VSMacros
{
    
    [ProvideToolWindow(typeof(MacrosToolWindow), Style = VsDockStyle.Tabbed, Window = "3ae79031-e1bc-11d0-8f78-00a0c9110057")]
    [ProvideKeyBindingTable(GuidList.GuidToolWindowPersistanceString, 115)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid(GuidList.GuidVSMacrosPkgString)]
    public sealed class VSMacrosPackage : Package
    {
        
        public static VSMacrosPackage Current { get; private set; }

        public VSMacrosPackage()
        {
            VSMacrosPackage.Current = this;
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

        protected override int QueryClose(out bool canClose)
        {
            return base.QueryClose(out canClose);
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

        public void PlaybackCommand1(object sender, EventArgs arguments) { Manager.Instance.PlaybackCommand(1); }
        public void PlaybackCommand2(object sender, EventArgs arguments) { Manager.Instance.PlaybackCommand(2); }
        public void PlaybackCommand3(object sender, EventArgs arguments) { Manager.Instance.PlaybackCommand(3); }
        public void PlaybackCommand4(object sender, EventArgs arguments) { Manager.Instance.PlaybackCommand(4); }
        public void PlaybackCommand5(object sender, EventArgs arguments) { Manager.Instance.PlaybackCommand(5); }
        public void PlaybackCommand6(object sender, EventArgs arguments) { Manager.Instance.PlaybackCommand(6); }
        public void PlaybackCommand7(object sender, EventArgs arguments) { Manager.Instance.PlaybackCommand(7); }
        public void PlaybackCommand8(object sender, EventArgs arguments) { Manager.Instance.PlaybackCommand(8); }
        public void PlaybackCommand9(object sender, EventArgs arguments) { Manager.Instance.PlaybackCommand(9); }

        #endregion
    }
}
