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
using VSMacros.Model;
using System.ComponentModel.Design; // for CommandID

namespace VSMacros
{
    [Guid("56fbfa32-c049-4fd5-9b54-39fcdf33629d")]
    public class ExplorerToolWindow : ToolWindowPane
    {
        public ExplorerToolWindow() :
            base(null)
        {
            this.Caption = Resources.ToolWindowTitle;
            this.BitmapResourceID = 301;
            this.BitmapIndex = 1;

            // Instantiate Tool Window Toolbar
            this.ToolBar = new CommandID(GuidList.guidVSMacrosCmdSet, PkgCmdIDList.TWToolbar);

            string MacroDirectory = VSMacrosPackage.Current.MacroDirectory;
            base.Content = new MacroBrowserList(new MacroFSNode(MacroDirectory));
        }
    }
}
