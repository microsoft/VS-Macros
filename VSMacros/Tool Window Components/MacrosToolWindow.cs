using System;
using System.Collections;
using System.ComponentModel;
using System.ComponentModel.Design; // for CommandID
using System.Data;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using VSMacros.Engines;
using VSMacros.Models;

namespace VSMacros
{
    [Guid("56fbfa32-c049-4fd5-9b54-39fcdf33629d")]
    public class MacrosToolWindow : ToolWindowPane
    {
        private const string currentMacroLocation = "Current.js";

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

            if (!File.Exists(Path.Combine(VSMacrosPackage.Current.MacroDirectory, currentMacroLocation)))
            {
                File.Create(Path.Combine(VSMacrosPackage.Current.MacroDirectory, currentMacroLocation));
            }

            string MacroDirectory = VSMacrosPackage.Current.MacroDirectory;

            // Create tree view root
            MacroFSNode root = new MacroFSNode(MacroDirectory);

            // Make sure it is opened by default
            root.IsExpanded = true;

            // Initialize Macros Control
            base.Content = new MacrosControl(root);
        }
    }
}
