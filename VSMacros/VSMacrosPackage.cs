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
        private BitmapImage startIcon;
        private BitmapImage stopIcon;
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
                   ShowToolWindow,
                   new CommandID(GuidList.GuidVSMacrosCmdSet, (int)PkgCmdIDList.CmdIdMacroExplorer)));

                // Create the command for start recording
                mcs.AddCommand(new MenuCommand(
                  Record,
                  new CommandID(GuidList.GuidVSMacrosCmdSet, (int)PkgCmdIDList.CmdIdRecord)));

                // Create the command for playback
                mcs.AddCommand(new MenuCommand(
                    Playback,
                    new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdPlayback)));

                // Create the command for playback multiple times
                mcs.AddCommand(new MenuCommand(
                    PlaybackMultipleTimes,
                    new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdPlaybackMultipleTimes)));

                // Create the command for save current macro
                mcs.AddCommand(new MenuCommand(
                    SaveCurrent,
                    new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdSaveTemporaryMacro)));

                // Create the command for refresh
                mcs.AddCommand(new MenuCommand(
                    Refresh,
                    new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdRefresh)));

                // Create the command to open the macro directory
                mcs.AddCommand(new MenuCommand(
                    OpenDirectory,
                    new CommandID(GuidList.GuidVSMacrosCmdSet, PkgCmdIDList.CmdIdOpenDirectory)));
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
                this.StatusBarChange(Resources.StatusBarRecordingText, 1, this.StopIcon);
                IRecorder recorder = (IRecorder)this.GetService(typeof(IRecorder));
                recorder.StartRecording();
            }
            else
            {
                this.StatusBarChange(Resources.StatusBarReadyText, 0, this.StartIcon);
                IRecorder recorder = (IRecorder)this.GetService(typeof(IRecorder));
                recorder.StopRecording();
            }
            return;
        }

        private void Playback(object sender, EventArgs arguments)
        {
            Manager.Instance.Playback("", 1);
        }

        private void PlaybackMultipleTimes(object sender, EventArgs arguments)
        {
            Manager.Instance.Playback("", 0);
        }

        private void SaveCurrent(object sender, EventArgs arguments)
        {
            Manager.Instance.SaveCurrent();
        }

        private void Refresh(object sender, EventArgs arguments)
        {

        }

        public void OpenDirectory(object sender, EventArgs arguments)
        {
            // Open the macro directory and let the user manage the macros
            System.Threading.Tasks.Task.Run(() => { System.Diagnostics.Process.Start(MacroDirectory); });
        }

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
                    this.startIcon = new BitmapImage(new Uri(@"Resources\RecordRound.png"));
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
                    this.stopIcon = new BitmapImage(new Uri(@"Resources\stopIcon.png"));
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
            return (int)VSConstants.S_OK;
        }

        #endregion
    }
}
