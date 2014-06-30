using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using VSMacros.Engines;
using System.Windows.Media.Imaging;
using System.Windows.Forms;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.CommandBars;
using System.Collections.Generic;
using EnvDTE;
using Microsoft.Internal.VisualStudio.PlatformUI;

namespace VSMacros
{
    [ProvideToolWindow(typeof(MacrosToolWindow), Style = VsDockStyle.Tabbed, Window = "3ae79031-e1bc-11d0-8f78-00a0c9110057")]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid(GuidList.GuidVSMacrosPkgString)]
    [ProvideService(typeof(IMacroRecorder))]
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
        private bool isShowingStartImage = true;
        private List<CommandBarButton> imageButtons;
        private IVsStatusbar statusBar;
        private object iconRecord = (short)Microsoft.VisualStudio.Shell.Interop.Constants.SBAI_Synch;
        private static DataModel dataModel;

        internal static DataModel DataModel
        {
            get
            {
                if (VSMacrosPackage.dataModel == null)
                {
                    VSMacrosPackage.dataModel = new DataModel();
                }

                return VSMacrosPackage.dataModel;
            }
        }

        protected override void Initialize()
        {
            base.Initialize();

            ((IServiceContainer)this).AddService(typeof(IMacroRecorder), (serviceContainer, type) => { return new MacroRecorder(this); }, true);
            // Add our command handlers for the menu
            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if ( null != mcs )
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
            //Manager.Instance.ToggleRecording();
            if (isShowingStartImage)
            {
                statusBar = (IVsStatusbar)GetService(typeof(SVsStatusbar));
                statusBar.Clear();
                statusBar.SetText("Recording...");
                statusBar.Animation(1, ref iconRecord);
                foreach (var button in ImageButtons)
                {
                    button.Picture = (stdole.StdPicture)ImageHelper.IPictureFromBitmapSource(StopIcon);
                }

                IMacroRecorder macroRecorder = (IMacroRecorder)this.GetService(typeof(IMacroRecorder));
                macroRecorder.StartRecording();


            }
            else
            {
                statusBar.Clear();
                statusBar.SetText("Ready");
                statusBar.Animation(0, ref iconRecord);
                foreach (var button in ImageButtons)
                {
                    button.Picture = (stdole.StdPicture)ImageHelper.IPictureFromBitmapSource(StartIcon);
                }

                IMacroRecorder macroRecorder = (IMacroRecorder)this.GetService(typeof(IMacroRecorder));
                macroRecorder.StopRecording();

            }
            isShowingStartImage = !isShowingStartImage;

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

        internal List<CommandBarButton> ImageButtons
        {
            get
            {
                if (this.imageButtons == null)
                {
                    this.imageButtons = GetImageButtons();
                }

                return this.imageButtons;
            }
        }

        private List<CommandBarButton> GetImageButtons()
        {
            List<CommandBarButton> buttons = new List<CommandBarButton>();

            DTE dte = (DTE)this.GetService(typeof(SDTE));
            CommandBar mainMenu = ((CommandBars)dte.CommandBars)["MenuBar"];
            CommandBarPopup viewMenu = (CommandBarPopup)mainMenu.Controls["Tools"];
            CommandBarPopup toolMenu = (CommandBarPopup)viewMenu.Controls["Macros"];
            CommandBarButton startButton = (CommandBarButton)toolMenu.Controls["Start/Stop Recording"];
            buttons.Add(startButton);

            return buttons;
        }

        private BitmapSource StartIcon
        {
            get
            {
                if (this.startIcon == null)
                {
                    this.startIcon = new BitmapImage(new Uri(Path.Combine(CommonPath, "RecordRound.png")));
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
                    this.stopIcon = new BitmapImage(new Uri(Path.Combine(CommonPath, "stopIcon.png")));
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
            if (!isShowingStartImage)
            {
                string message = "Recording in process, are you sure to close the window?";
                string caption = "Attention";
                MessageBoxButtons buttons = MessageBoxButtons.YesNo;
                System.Windows.Forms.DialogResult result;

                // Displays the MessageBox.
                result = MessageBox.Show(message, caption, buttons);

                if (result == System.Windows.Forms.DialogResult.Yes)
                {
                    canClose = true;
                }
                else
                {
                    canClose = false;
                }
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
