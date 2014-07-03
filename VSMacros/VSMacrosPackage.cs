//-----------------------------------------------------------------------
// <copyright file="VSMacrosPackage.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Drawing;
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
    [ProvideToolWindow(typeof(MacrosToolWindow), Style = VsDockStyle.Tabbed, Window = "3ae79031-e1bc-11d0-8f78-00a0c9110057")]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid(GuidList.GuidVSMacrosPkgString)]
    [ProvideService(typeof(IRecorder))]
    public sealed class VSMacrosPackage : Package
    {
        public static VSMacrosPackage Current { get; private set; }

        public VSMacrosPackage()
        {
            Current = this;
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
                    this.macroDirectory = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Macros");
                return this.macroDirectory;
            }
        }

        /////////////////////////////////////////////////////////////////////////////
        // Overridden Package Implementation
        #region Package Members
        private BitmapSource startIcon;
        private BitmapSource stopIcon;
        private string commonPath;
        private List<CommandBarButton> imageButtons;
        private IVsStatusbar statusBar;
        private object iconRecord = (short)Microsoft.VisualStudio.Shell.Interop.Constants.SBAI_Synch;
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

        public void Record(object sender, EventArgs arguments)
        {
            IRecorderPrivate macroRecorder = (IRecorderPrivate)this.GetService(typeof(IRecorder));
            if (!macroRecorder.IsRecording)
            {
                this.StatusBarChange(Resources.StatusBarRecordingText, 1, this.StartIcon);
                Manager.Instance.StartRecording();
            }
            else
            {
                this.StatusBarChange(Resources.StatusBarReadyText, 0, this.StopIcon);
                Manager.Instance.StopRecording();
            }
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

        private void StatusBarChange(string status, int animation, BitmapSource icon)
        {
            this.statusBar.Clear();
            this.statusBar.SetText(status);
            this.statusBar.Animation(animation, ref this.iconRecord);

            foreach (CommandBarButton button in this.ImageButtons)
            {
                try
                {
                    if (button != null)
                    {
                        button.Picture = (stdole.StdPicture)ImageHelper.IPictureFromBitmapSource(icon);
                    }
                }
                catch (Exception menuButtonRemoved)
                {
                    // Do nothing since the removed button does not need to change its image;
                }
            }
        }

        internal List<CommandBarButton> ImageButtons
        {
            get
            {
                if (this.imageButtons == null)
                {
                    List<CommandBarButton> buttons = new List<CommandBarButton>();
                    this.imageButtons = this.AddMenuButton(buttons);
                }
                return this.imageButtons;
            }
        }

        private List<CommandBarButton> AddMenuButton(List<CommandBarButton> buttons)
        {
            List<CommandBarButton> buttonsList = buttons;
            DTE dte = (DTE)this.GetService(typeof(SDTE));
            CommandBar mainMenu = ((CommandBars)dte.CommandBars)["MenuBar"];
            CommandBarPopup toolMenu = (CommandBarPopup)mainMenu.Controls["Tools"];
            CommandBarPopup macroMenu = (CommandBarPopup)toolMenu.Controls["Macros"];
            if (macroMenu != null)
            {
                CommandBarButton startButton = (CommandBarButton)macroMenu.Controls["Start/Stop Recording"];
                if (startButton != null)
                {
                    buttons.Add(startButton);
                }
            }
            return buttonsList;
        }

        private BitmapSource StartIcon
        {
            get
            {
                if (this.startIcon == null)
                {
                    Bitmap bmp = Resources.RecordRound;

                    this.startIcon = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                            bmp.GetHbitmap(),
                            IntPtr.Zero,
                            System.Windows.Int32Rect.Empty,
                            BitmapSizeOptions.FromEmptyOptions());
                }

                return this.startIcon;
            }
        }

        private BitmapSource StopIcon
        {
            get
            {
                if (this.stopIcon == null)
                {
                    Bitmap bmp = Resources.StopIcon;

                    this.startIcon = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                            bmp.GetHbitmap(),
                            IntPtr.Zero,
                            System.Windows.Int32Rect.Empty,
                            BitmapSizeOptions.FromEmptyOptions());
                }
                return this.stopIcon;
            }
        }

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
