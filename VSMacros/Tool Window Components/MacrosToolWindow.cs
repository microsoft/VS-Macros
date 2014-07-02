//-----------------------------------------------------------------------
// <copyright file="MacrosToolWindow.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections;
using System.ComponentModel;
using System.ComponentModel.Design; // for CommandID
using System.Data;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using VSMacros.Engines;
using VSMacros.Model;
using System.Reflection;

namespace VSMacros
{
    [Guid("56fbfa32-c049-4fd5-9b54-39fcdf33629d")]
    public class MacrosToolWindow : ToolWindowPane
    {
        private const string currentMacroLocation = "Current.js";
        private VSMacrosPackage owningPackage;
        private bool addedToolbarButton;

        public MacrosToolWindow() :
            base(null)
        {
            this.owningPackage = VSMacrosPackage.Current;
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
            MacroFSNode root = new MacroFSNode(MacroDirectory);
            var macroControl = new MacrosControl(root);
            macroControl.Loaded += this.OnLoaded;
            base.Content = macroControl;
        }

        public void OnLoaded(object sender, System.Windows.RoutedEventArgs e)
        {
            if (!this.addedToolbarButton)
            {
                IVsWindowFrame windowFrame = (IVsWindowFrame)GetService(typeof(SVsWindowFrame));

                object dteWindow;
                windowFrame.GetProperty((int)__VSFPROPID.VSFPROPID_ExtWindowObject, out dteWindow);
                Window2 window = (Window2)dteWindow;
                this.owningPackage.ImageButtons.Add((CommandBarButton)((CommandBars)window.CommandBars)[1].Controls[1]);
                this.addedToolbarButton = true;
            }
        }
    }
}
