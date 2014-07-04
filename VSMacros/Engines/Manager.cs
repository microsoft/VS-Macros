using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using VSMacros.Dialogs;
using VSMacros.Interfaces;
using VSMacros.Models;
using System.Diagnostics;

namespace VSMacros.Engines
{
    internal sealed class Manager : IManager
    {
        private static readonly Manager instance = new Manager();

        public string[] Shortcuts { get; private set; }
        private bool shortcutsLoaded;
        private const string shortcutsFileName = "Shortcuts.xml";
        private string shortcutsFilePath;

        private IVsUIShell uiShell;
        private bool uiShellLoaded;

        private MacroFSNode SelectedMacro
        {
            get { return MacrosControl.Current.SelectedNode; }
        }

        private Manager() 
        {
            uiShell = (IVsUIShell)((IServiceProvider)VSMacrosPackage.Current).GetService(typeof(SVsUIShell));
            if (uiShell != null)
            {
                uiShellLoaded = true;
            }
            else
            {
                uiShellLoaded = false;
            }

            this.shortcutsFilePath = Path.Combine(VSMacrosPackage.Current.MacroDirectory, shortcutsFileName);
            this.LoadShortcuts();
            shortcutsLoaded = true;
        }

        public static Manager Instance
        {
            get { return instance; }
        }

        public void ToggleRecording()
        {
        }

        public void Playback(string path, int times) 
        {
            if (path == string.Empty)
            {
                path = this.SelectedMacro.FullPath;
            }

            var executor = new Executor();

            if (!Executor.IsEngineInitialized)
            {
                executor.InitializeEngine();
                // TODO: eventually we'll want to run the macro here as well
            }
            else
            {
                executor.Shenanigans(path);
            }
        }

        public void StopPlayback() 
        {
        }

        public void OpenFolder(string path = null)
        {
            if (path == null)
            {
                path = SelectedMacro.FullPath;
            }

            // Open the macro directory and let the user manage the macros
            System.Threading.Tasks.Task.Run(() => { System.Diagnostics.Process.Start(path); });
        }

        public void SaveCurrent() 
        {
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.InitialDirectory = VSMacrosPackage.Current.MacroDirectory;
            dlg.FileName = "My Macro";
            dlg.DefaultExt = ".js";
            dlg.Filter = "Macro| *.js";

            Nullable<bool> result = dlg.ShowDialog();

            /* QUESTION should I use IVsUIShell.GetSaveFileNameViaDlg?
            if (uiShellLoaded)
            {
                VSSAVEFILENAMEW[] saveFileNameInfo = new VSSAVEFILENAMEW[1];

                saveFileNameInfo[0].pwzFileName = "My Macro";
                saveFileNameInfo[0].nFileExtension = ".js";
                saveFileNameInfo[0].pwzFilter = "Macro| *.js";

                uiShell.GetSaveFileNameViaDlg(saveFileNameInfo);
            }
            */

            if (result == true)
            {
                 string pathToNew = dlg.FileName;

                try
                {
                    string pathToCurrent = Path.Combine(VSMacrosPackage.Current.MacroDirectory, "Current.js");

                    File.Copy(pathToCurrent, pathToNew);

                    this.Refresh();
                }
                catch (Exception e)
                {
                    this.ShowMessageBox(e.Message);
                }
            }
        }

        public void Refresh()
        {
            MacroFSNode.RootNode.RefreshTree();
            this.LoadShortcuts();
        }

        public void Edit()
        {
            MacroFSNode macro = this.SelectedMacro;
            string path = macro.FullPath;

            // QUESTION Do I need to use the method 'OpenDocument(IServiceProvider, String, Guid, IVsUIHierarchy, UInt32, IVsWindowFrame)' instead? 
            VsShellUtilities.OpenDocument(VSMacrosPackage.Current, path);
        }

        public void Rename()
        {
            MacroFSNode macro = this.SelectedMacro;
            macro.EnableEdit();
        }

        public void AssignShortcut()
        {
            AssignShortcutDialog dlg = new AssignShortcutDialog();
            dlg.ShowDialog();

            if (dlg.DialogResult == true)
            {
                MacroFSNode macro = this.SelectedMacro;

                // Remove old shortcut if it exists
                int oldKey;
                if (macro.Shortcut != string.Empty)
                {
                    string oldShortcut = macro.Shortcut;
                    oldKey = (int) oldShortcut[oldShortcut.Length - 2] - '0';
                    this.Shortcuts[oldKey] = string.Empty;
                }

                int newShortcutNumber = dlg.SelectedShortcutNumber;

                // At this point, the shortcut has been removed
                // Assign a new one only if the user selected a key binding
                if (newShortcutNumber != 0)
                {
                    // Update dictionary
                    this.Shortcuts[newShortcutNumber] = macro.FullPath;
                }

                // Notify the change
                if (dlg.shouldRefreshFileSystem)
                {
                    // Refresh entire fs
                    this.SaveShortcuts();
                    this.Refresh();
                }
                else
                {
                    // Refresh selected macro
                    macro.Shortcut = string.Empty;
                }
            }
        }

        public void Delete()
        {
            MacroFSNode macro = this.SelectedMacro;
            string path = macro.FullPath;

            FileSystemInfo file;
            string fileName = Path.GetFileNameWithoutExtension(path);
            string message;

            if (macro.IsDirectory)
            {
                file = new DirectoryInfo(path);
                message = string.Format(VSMacros.Resources.DeleteFolder, fileName);
            }
            else
            {
                file = new FileInfo(path);
                message = string.Format(VSMacros.Resources.DeleteMacro, fileName);
            }

            if (file.Exists)
            {
                int result;
                result = this.ShowMessageBox(message, OLEMSGBUTTON.OLEMSGBUTTON_OKCANCEL);

                // TODO 1 is for OK (replace 1 by IDOK)
                if (result == 1)
                {
                    try
                    {
                        file.Delete();  // TODO non-empty dir will raise an exception here -> User Directory.Delete(path, true)
                        macro.Delete();
                    }
                    catch (Exception e)
                    {
                        this.ShowMessageBox(e.Message);
                    }
                }
            }
            else
            {
                macro.Delete();
            }
        }
      
        public void PlaybackCommand(int cmd)
        {
            // Load shortcuts if not already loaded
            if (!this.shortcutsLoaded)
            {
                this.LoadShortcuts();
            }
            
            // Get path to macro bound to the shortcut
            string path = this.Shortcuts[cmd];

            if (path != string.Empty)
            {
                this.Playback(path, 1);
            }
        }

        private StreamReader LoadFile(string path) 
        { 
            try
            {
                if (!File.Exists(path))
                {
                    throw new Exception(Resources.MacroNotFound);
                }

                StreamReader str = new StreamReader(path);
                
                return str;
            }
            catch (Exception e)
            {
                this.ShowMessageBox(e.Message);
            }

            return null;
        }

        private void SaveMacro(Stream str, string path)
        { 
            try
            {
                using (var fileStream = File.Create(path))
                {
                    str.Seek(0, SeekOrigin.Begin);
                    str.CopyTo(fileStream);
                }
            }
            catch (Exception e)
            {
                this.ShowMessageBox(e.Message);
            }
        }

        private void LoadShortcuts()
        {
            shortcutsLoaded = true;

            try
            {
                // Load XML file
                var root = XDocument.Load(Path.Combine(VSMacrosPackage.Current.MacroDirectory, shortcutsFilePath));

            // Parse to dictionary
            this.Shortcuts = root.Descendants("command")
                                            .Select(elmt => elmt.Value)
                                            .ToArray();
            }
            catch (Exception e)
            {
                this.ShowMessageBox(e.Message);
            }
        }

        private void SaveShortcuts()
        {
            XDocument xmlShortcuts = 
                new XDocument(
                    new XDeclaration("1.0", "utf-8", "yes"),
                    new XElement("commands",
                        from s in this.Shortcuts
                            select new XElement("command",
                                new XText(s))));

            xmlShortcuts.Save(shortcutsFilePath);
        }

        #region Helper Methods
        public int ShowMessageBox(string message, OLEMSGBUTTON btn = OLEMSGBUTTON.OLEMSGBUTTON_OK)
        {
            if (!uiShellLoaded)
            {
                return -1;
            }

            Guid clsid = Guid.Empty;
            int result;

            Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(
              uiShell.ShowMessageBox(
                0,
                ref clsid,
                string.Empty,
                message,
                string.Empty,
                0,
                btn,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST,
                OLEMSGICON.OLEMSGICON_WARNING,
                0,        // false
                out result));

            return result;
        }
        #endregion
    }
}
