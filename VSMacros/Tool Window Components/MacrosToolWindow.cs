using System;
using System.IO;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Shell;
using System.ComponentModel.Design; // for CommandID

using VSMacros.Model;
using VSMacros.Engines;


namespace VSMacros
{
    [Guid("56fbfa32-c049-4fd5-9b54-39fcdf33629d")]
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

            // Initialize file system
            if (!Directory.Exists(VSMacrosPackage.Current.MacroDirectory))
                Directory.CreateDirectory(VSMacrosPackage.Current.MacroDirectory);

            if (!File.Exists(Path.Combine(VSMacrosPackage.Current.MacroDirectory, "Current.js")))
                File.Create(Path.Combine(VSMacrosPackage.Current.MacroDirectory, "Current.js"));

            // Load Current macro
            Manager.Instance.LoadCurrent();

            string MacroDirectory = VSMacrosPackage.Current.MacroDirectory;
            base.Content = new MacrosControl(new MacroFSNode(MacroDirectory));
        }
    }
}
