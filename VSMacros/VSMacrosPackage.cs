//-----------------------------------------------------------------------
// <copyright file="VSMacrosPackage.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Media.Imaging;
using EnvDTE;
using Microsoft.Internal.VisualStudio.PlatformUI;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using VSMacros.Engines;
using VSMacros.Interfaces;
using VSMacros.Model;

namespace VSMacros
{
    [Guid(GuidList.GuidVSMacrosPkgString)]
    [ProvideToolWindow(typeof(MacrosToolWindow), Style = VsDockStyle.Tabbed, Window = "3ae79031-e1bc-11d0-8f78-00a0c9110057")]
    public sealed class VSMacrosPackage : Package
    {
        private static VSMacrosPackage current;
        public static VSMacrosPackage Current
        {
            get
            {
                if (current == null)
                {
                    current = new VSMacrosPackage();
                }

                return current;
            }
        }

        public VSMacrosPackage()
        {
            VSMacrosPackage.current = this;
        }

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

        private string macroDirectory;
        public string MacroDirectory
        {
            get
            {
                if (this.macroDirectory == default(string))
                {
                    this.macroDirectory = Path.Combine(this.UserLocalDataPath, "Macros");
                }
                return this.macroDirectory;
            }
        }

        private string assemblyDirectory;
        public string AssemblyDirectory
        {
            get
            {
                if (this.assemblyDirectory == default(string))
                {
                    this.assemblyDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                }

                return this.assemblyDirectory;
            }
        }

        /////////////////////////////////////////////////////////////////////////////
        // Overridden Package Implementation
        #region Package Members
        private BitmapImage startIcon;
        private BitmapImage playbackIcon;
        private BitmapImage stopIcon;
        private string commonPath;
        private List<CommandBarButton> imageButtons;
        private IVsStatusbar statusBar;
        private object iconRecord = (short)Microsoft.VisualStudio.Shell.Interop.Constants.SBAI_General;
        private static RecorderDataModel dataModel;

        internal static RecorderDataModel DataModel
        {
            get
            {
                if (VSMacrosPackage.dataModel == null)
                {
                    VSMacrosPackage.dataModel = new RecorderDataModel();
                }

                return VSMacrosPackage.dataModel;
            }
        }

        protected override void Initialize()
        {
            base.Initialize();

            ((IServiceContainer)this).AddService(typeof(IRecorder), (serviceContainer, type) => { return new Recorder(this); }, promote: true);
            this.statusBar = (IVsStatusbar)GetService(typeof(SVsStatusbar));

            // Add our command handlers for the menu
            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs)
            {
                // Create the command for the tool window
                mcs.AddCommand(new MenuCommand(
                   this.ShowToolWindow,
                   new CommandID(GuidList.GuidVSMacrosCmdSet, (int)PkgCmdIDList.CmdIdMacroExplorer)));

                // Create the command for start recording
                CommandID recordCommandID = new CommandID(GuidList.GuidVSMacrosCmdSet, (int)PkgCmdIDList.CmdIdRecord);
                OleMenuCommand recordMenuItem = new OleMenuCommand(this.Record, recordCommandID);
                recordMenuItem.BeforeQueryStatus += new EventHandler(this.Record_OnBeforeQueryStatus);
                mcs.AddCommand(recordMenuItem);

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
        #endregion

        /////////////////////////////////////////////////////////////////////////////
        // Command Handlers
        #region Command Handlers

        private void Record(object sender, EventArgs arguments)
        {
            IRecorderPrivate macroRecorder = (IRecorderPrivate)this.GetService(typeof(IRecorder));
            if (!macroRecorder.IsRecording)
            {
                Manager.Instance.StartRecording();

                this.StatusBarChange(Resources.StatusBarRecordingText, 1);
                this.ChangeMenuIcons(this.StopIcon, 0);
                this.UpdateButtonsForRecording(true);
            }
            else
            {
                Manager.Instance.StopRecording();

                this.StatusBarChange(Resources.StatusBarReadyText, 0);
                this.ChangeMenuIcons(this.StartIcon, 0);
                this.UpdateButtonsForRecording(false);
            }
        }

        public void Playback(object sender, EventArgs arguments)
        {
            if (Manager.Instance.executor == null || !Manager.Instance.executor.IsEngineRunning)
            {
                //this.UpdateButtonsForPlayback(true);
            }
            else
            {
                this.UpdateButtonsForPlayback(false);
            }

            Manager.Instance.Playback(string.Empty);
        }

        private void PlaybackMultipleTimes(object sender, EventArgs arguments)
        {
            if (Manager.Instance.executor == null || !Manager.Instance.executor.IsEngineRunning)
            {
                //this.UpdateButtonsForPlayback(true);
            }
            else
            {
                this.UpdateButtonsForPlayback(false);
            }

            Manager.Instance.PlaybackMultipleTimes(string.Empty);
        }

        private void SaveCurrent(object sender, EventArgs arguments)
        {
            Manager.Instance.SaveCurrent();
        }

        private void PlaybackCommand1(object sender, EventArgs arguments) { Manager.Instance.PlaybackCommand(1); }
        private void PlaybackCommand2(object sender, EventArgs arguments) { Manager.Instance.PlaybackCommand(2); }
        private void PlaybackCommand3(object sender, EventArgs arguments) { Manager.Instance.PlaybackCommand(3); }
        private void PlaybackCommand4(object sender, EventArgs arguments) { Manager.Instance.PlaybackCommand(4); }
        private void PlaybackCommand5(object sender, EventArgs arguments) { Manager.Instance.PlaybackCommand(5); }
        private void PlaybackCommand6(object sender, EventArgs arguments) { Manager.Instance.PlaybackCommand(6); }
        private void PlaybackCommand7(object sender, EventArgs arguments) { Manager.Instance.PlaybackCommand(7); }
        private void PlaybackCommand8(object sender, EventArgs arguments) { Manager.Instance.PlaybackCommand(8); }
        private void PlaybackCommand9(object sender, EventArgs arguments) { Manager.Instance.PlaybackCommand(9); }

        #endregion

        #region Status Bar & Menu Icons
        public void ChangeMenuIcons(BitmapSource icon, int commandNumber)
        {
            // commandNumber is 0 for Recording, 1 for Playback and 2 for Playback Multiple Times         
            try
            {
                if (this.ImageButtons[commandNumber] != null)
                {
                    // Change icon in menu
                    this.ImageButtons[commandNumber].Picture = (stdole.StdPicture)ImageHelper.IPictureFromBitmapSource(icon);

                    if (this.ImageButtons.Count > 3)
                    {
                        // Change icon in toolbar
                        this.ImageButtons[commandNumber + 3].Picture = (stdole.StdPicture)ImageHelper.IPictureFromBitmapSource(icon);
                    }
                }
            }
            catch (ObjectDisposedException)
            {
                // Do nothing since the removed button does not need to change its image;
            }
        }

        internal void StatusBarChange(string status, int animation)
        {
            this.statusBar.Clear();
            this.statusBar.SetText(status);
        }

        internal List<CommandBarButton> ImageButtons
        {
            get
            {
                if (this.imageButtons == null)
                {
                    this.imageButtons = new List<CommandBarButton>();
                    this.AddMenuButton();
                }
                return this.imageButtons;
            }
        }

        private void AddMenuButton()
        {
            DTE dte = (DTE)this.GetService(typeof(SDTE));
            CommandBar mainMenu = ((CommandBars)dte.CommandBars)["MenuBar"];
            CommandBarPopup toolMenu = (CommandBarPopup)mainMenu.Controls["Tools"];
            CommandBarPopup macroMenu = (CommandBarPopup)toolMenu.Controls["Macros"];
            if (macroMenu != null)
            {
                try
                {
                    List<CommandBarButton> buttons = new List<CommandBarButton>()
                    {
                        (CommandBarButton)macroMenu.Controls["Start Recording"],
                        (CommandBarButton)macroMenu.Controls["Playback"],
                        (CommandBarButton)macroMenu.Controls["Playback Multiple Times"]
                    };

                    foreach (var item in buttons)
                    {
                        if (item != null)
                        {
                            this.imageButtons.Add(item);
                        }
                    }
                }
                catch (ArgumentException)
                {
                    // nothing to do
                }
            }
        }

        private void UpdateButtonsForRecording(bool isRecording)
        {
            this.EnableMyCommand(PkgCmdIDList.CmdIdPlayback, !isRecording);
            this.EnableMyCommand(PkgCmdIDList.CmdIdPlaybackMultipleTimes, !isRecording);
            this.UpdateCommonButtons(!isRecording);
        }

        public void UpdateButtonsForPlayback(bool goingToPlay)
        {
            this.EnableMyCommand(PkgCmdIDList.CmdIdRecord, !goingToPlay);
            this.EnableMyCommand(PkgCmdIDList.CmdIdPlaybackMultipleTimes, !goingToPlay);
            this.UpdateCommonButtons(!goingToPlay);

            if (goingToPlay)
            {
                this.ChangeMenuIcons(this.StopIcon, 1);
            }
            else
            {
                this.ChangeMenuIcons(this.PlaybackIcon, 1);
            }
        }

        private void UpdateCommonButtons(bool enable)
        {
            this.EnableMyCommand(PkgCmdIDList.CmdIdSaveTemporaryMacro, enable);
            this.EnableMyCommand(PkgCmdIDList.CmdIdRefresh, enable);
            this.EnableMyCommand(PkgCmdIDList.CmdIdOpenDirectory, enable);
        }

        internal bool EnableMyCommand(int cmdID, bool enableCmd)
        {
            bool cmdUpdated = false;
            var mcs = this.GetService(typeof(IMenuCommandService))
                    as OleMenuCommandService;
            var newCmdID = new CommandID(GuidList.GuidVSMacrosCmdSet, cmdID);
            MenuCommand mc = mcs.FindCommand(newCmdID);
            if (mc != null)
            {
                mc.Enabled = enableCmd;
                cmdUpdated = true;
            }
            return cmdUpdated;
        }

        internal void ClearStatusBar()
        {
            this.StatusBarChange(Resources.StatusBarReadyText, 0);
        }

        private BitmapSource StartIcon
        {
            get
            {
                if (this.startIcon == null)
                {
                    this.startIcon = new BitmapImage(new Uri(Path.Combine(this.CommonPath, "RecordRound.png")));
                }
                return this.startIcon;
            }
        }

        private BitmapSource PlaybackIcon
        {
            get
            {
                if (this.playbackIcon == null)
                {
                    this.playbackIcon = new BitmapImage(new Uri(Path.Combine(this.CommonPath, "PlaybackIcon.png")));
                }
                return this.playbackIcon;
            }
        }

        internal BitmapSource StopIcon
        {
            get
            {
                if (this.stopIcon == null)
                {
                    this.stopIcon = new BitmapImage(new Uri(Path.Combine(this.CommonPath, "StopIcon.png")));
                }
                return this.stopIcon;
            }
        }

        #endregion

        private string CommonPath
        {
            get
            {
                if (this.commonPath == null)
                {
                    this.commonPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Resources");
                }
                return this.commonPath;
            }
        }

        protected override int QueryClose(out bool canClose)
        {
            IRecorderPrivate macroRecorder = (IRecorderPrivate)this.GetService(typeof(IRecorder));
            if (macroRecorder.IsRecording)
            {
                string message = Resources.ExitMessage;
                string caption = Resources.ExitCaption;
                MessageBoxButtons buttons = MessageBoxButtons.YesNo;
                System.Windows.Forms.DialogResult result;

                // Displays the MessageBox.
                result = MessageBox.Show(message, caption, buttons);
                canClose = (result == System.Windows.Forms.DialogResult.Yes);
            }
            else
            {
                canClose = true;
            }

            // Close manager
            Manager.Instance.Close();

            if (Executor.Job != null)
            {
                Executor.Job.Close();
            }

            return (int)VSConstants.S_OK;
        }

        private void Record_OnBeforeQueryStatus(object sender, EventArgs e)
        {
            var recordCommand = sender as OleMenuCommand;

            if (recordCommand != null)
            {
                IRecorderPrivate macroRecorder = (IRecorderPrivate)this.GetService(typeof(IRecorder));
                if (macroRecorder.IsRecording)
                {
                    recordCommand.Text = Resources.MenuTextRecording;
                }
                else
                {
                    recordCommand.Text = Resources.MenuTextNormal;
                }
            }
        }
    }
}
