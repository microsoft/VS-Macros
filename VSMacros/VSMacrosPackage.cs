using System;
using System.IO;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using Microsoft.Win32;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using VSMacros.Engines;

using EnvDTE;

namespace VSMacros
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    ///
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the 
    /// IVsPackage interface and uses the registration attributes defined in the framework to 
    /// register itself and its components with the shell.
    /// </summary>
    
    [ProvideToolWindow(typeof(ExplorerToolWindow), Style = VsDockStyle.Tabbed, Window = "3ae79031-e1bc-11d0-8f78-00a0c9110057")]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid(GuidList.guidVSMacrosPkgString)]
    public sealed class VSMacrosPackage : Package
    {
        public static VSMacrosPackage Current { get; private set; }

        /// <summary>
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require 
        /// any Visual Studio service because at this point the package object is created but 
        /// not sited yet inside Visual Studio environment. The place to do all the other 
        /// initialization is the Initialize method.
        /// </summary>
        public VSMacrosPackage()
        {
            Current = this;
        }

        /// <summary>
        /// This function is called when the user clicks the menu item that shows the 
        /// tool window. See the Initialize method to see how the menu item is associated to 
        /// this function using the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        private void ShowToolWindow(object sender = null, EventArgs e = null)
        {
            // Get the (only) instance of this tool window
            // The last flag is set to true so that if the tool window does not exists it will be created.
            ToolWindowPane window = this.FindToolWindow(typeof(ExplorerToolWindow), 0, true);
            if ((null == window) || (null == window.Frame))
            {
                throw new NotSupportedException(Resources.CanNotCreateWindow);
            }
            IVsWindowFrame windowFrame = (IVsWindowFrame)window.Frame;
            Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(windowFrame.Show());
        }

        public string MacroDirectory
        {
            get { return Path.Combine(this.UserLocalDataPath, "Macros");  }
        }

        /////////////////////////////////////////////////////////////////////////////
        // Overridden Package Implementation
        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            base.Initialize();

            // Initialize file system
            if (!Directory.Exists(MacroDirectory))
                Directory.CreateDirectory(MacroDirectory);

            if (!File.Exists(Path.Combine(MacroDirectory, "Current.js")))
                File.Create(Path.Combine(MacroDirectory, "Current.js"));

            // Load Current macro
            Manager.Instance.LoadCurrent();

            // Add our command handlers for the menu
            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if ( null != mcs )
            {
                // Create the command for the tool window
                mcs.AddCommand(new MenuCommand(
                   ShowToolWindow,
                   new CommandID(GuidList.guidVSMacrosCmdSet, (int)PkgCmdIDList.cmdidMacroExplorer)));

                 // Create the command for start recording
                mcs.AddCommand(new MenuCommand(
                    Record,
                    new CommandID(GuidList.guidVSMacrosCmdSet, PkgCmdIDList.cmdidRecord)));

                // Create the command for playbback
                mcs.AddCommand(new MenuCommand(
                    Playback,
                    new CommandID(GuidList.guidVSMacrosCmdSet, PkgCmdIDList.cmdidPlayback)));

                // Create the command for playback multiple times
                mcs.AddCommand(new MenuCommand(
                    PlaybackMultipleTimes,
                    new CommandID(GuidList.guidVSMacrosCmdSet, PkgCmdIDList.cmdidPlaybackMultipleTimes)));

                // Create the command for save current macro
                mcs.AddCommand(new MenuCommand(
                    SaveCurrent,
                    new CommandID(GuidList.guidVSMacrosCmdSet, PkgCmdIDList.cmdidSaveTemporaryMacro)));
                // Create the command for refresh
                mcs.AddCommand(new MenuCommand(
                    Refresh,
                    new CommandID(GuidList.guidVSMacrosCmdSet, PkgCmdIDList.cmdidRefresh)));

                // Create the command to open the macro directory
                mcs.AddCommand(new MenuCommand(
                    OpenDirectory,
                    new CommandID(GuidList.guidVSMacrosCmdSet, PkgCmdIDList.cmdidOpenDirectory)));
            }
        }
        #endregion

        /////////////////////////////////////////////////////////////////////////////
        // Command Handlers
        #region Command Handlers

        public void Record(object sender, EventArgs arguments)
        {
            Manager.Instance.StartRecording();
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
            System.Diagnostics.Process.Start(MacroDirectory);
        }

        #endregion

    }
}
