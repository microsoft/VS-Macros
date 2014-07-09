//-----------------------------------------------------------------------
// <copyright file="Manager.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using VSMacros.Dialogs;
using VSMacros.Interfaces;
using VSMacros.Models;

namespace VSMacros.Engines
{
    public sealed class Manager : IManager
    {
        private static Manager instance;
        
        private const string CurrentMacroFileName = "Current.js";
        private const string ShortcutsFileName = "Shortcuts.xml";

        public static string CurrentMacroPath = Path.Combine(VSMacrosPackage.Current.MacroDirectory, Manager.CurrentMacroFileName);
        private static string dteIntellisensePath = Path.Combine(VSMacrosPackage.Current.AssemblyDirectory, "Intellisense", "dte.js");

        public string[] Shortcuts { get; private set; }
        private bool shortcutsLoaded;
       
        private string shortcutsFilePath;

        private IServiceProvider serviceProvider;
        private IVsUIShell uiShell;
        private bool uiShellLoaded;

        private MacroFSNode SelectedMacro
        {
            get { return MacrosControl.Current.SelectedNode; }
        }

        private Manager(IServiceProvider provider) 
        {
            this.serviceProvider = provider;
            this.uiShell = (IVsUIShell)provider.GetService(typeof(SVsUIShell));
            if (this.uiShell != null)
            {
                this.uiShellLoaded = true;
            }
            else
            {
                this.uiShellLoaded = false;
            }

            this.shortcutsFilePath = Path.Combine(VSMacrosPackage.Current.MacroDirectory, Manager.ShortcutsFileName);
            this.LoadShortcuts();
            this.shortcutsLoaded = true;
        }

        public static Manager Instance
        {
            get
            {
                if (Manager.instance == null)
                {
                    Manager.instance = new Manager(VSMacrosPackage.Current);
                }
                
                return Manager.instance;
            }
        }

        public void StartRecording()
        {
            IRecorder recorder = (IRecorder)this.serviceProvider.GetService(typeof(IRecorder));
            recorder.StartRecording();
        }

        public void StopRecording()
        {
            string current = Manager.CurrentMacroPath;

            IRecorder recorder = (IRecorder)this.serviceProvider.GetService(typeof(IRecorder));
            recorder.StopRecording(current);
        }

        public void Playback(string path, int times) 
        {
            if (path == string.Empty)
            {
                path = this.SelectedMacro.FullPath;
            }

            Executor executor = new Executor();
            executor.StartExecution(new StreamReader(path), 1);
        }

        public void StopPlayback() 
        {
        }

        public void OpenFolder(string path = null)
        {
            if (path == null)
            {
                path = this.SelectedMacro.FullPath;
            }

            // Open the macro directory and let the user manage the macros
            System.Threading.Tasks.Task.Run(() => { System.Diagnostics.Process.Start(path); });
        }

        public void SaveCurrent() 
        {
            SaveCurrentDialog dlg = new SaveCurrentDialog();
            bool? result = dlg.ShowDialog();

            if (result == true)
            {
                try
                {
                    string pathToNew = Path.Combine(VSMacrosPackage.Current.MacroDirectory, dlg.MacroName.Text + ".js");
                    string pathToCurrent = Manager.CurrentMacroPath;

                    int newShortcutNumber = dlg.SelectedShortcutNumber;

                    File.Copy(pathToCurrent, pathToNew);

                    MacroFSNode macro = new MacroFSNode(pathToNew, MacroFSNode.RootNode);

                    if (newShortcutNumber != 0)
                    {
                        // Update dictionary
                        this.Shortcuts[newShortcutNumber] = macro.FullPath;
                    }

                    // Notify the change
                    if (dlg.ShouldRefreshFileSystem)
                    {
                        // Notify all node that their shortcut property might have changed
                        this.SaveShortcuts();
                        MacroFSNode.NotifyAllNode(MacroFSNode.RootNode, "Shortcut");
                    }

                    this.Refresh();
                    MacroFSNode.FindNodeFromFullPath(pathToNew).IsSelected = true;
                }
                catch (Exception e)
                {
                    if (ErrorHandler.IsCriticalException(e)) { throw e; }
                    this.ShowMessageBox(e.Message);
                }
            }
        }

        public void Refresh()
        {
            this.SaveShortcuts();

            this.CreateFileSystem();

            MacroFSNode.RefreshTree();

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
                if (macro.Shortcut != MacroFSNode.NONE)
                {
                    this.Shortcuts[macro.Shortcut] = string.Empty;
                }

                int newShortcutNumber = dlg.SelectedShortcutNumber;

                // At this point, the shortcut has been removed
                // Assign a new one only if the user selected a key binding
                if (newShortcutNumber != MacroFSNode.NONE)
                {
                    // Update dictionary
                    this.Shortcuts[newShortcutNumber] = macro.FullPath;
                }

                // Notify the change
                if (dlg.ShouldRefreshFileSystem)
                {
                    // Notify all node that their shortcut property might have changed
                    this.SaveShortcuts();
                    this.Refresh();
                }
                else
                {
                    // Refresh selected macro
                    macro.Shortcut = MacroFSNode.TO_FETCH;
                }
            }
        }

        public void Delete()
        {
            MacroFSNode macro = this.SelectedMacro;

            // Don't delete if macro is being edited
            if (macro.IsEditable) { return; }

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
                VSConstants.MessageBoxResult result;
                result = this.ShowMessageBox(message, OLEMSGBUTTON.OLEMSGBUTTON_OKCANCEL);

                // TODO 1 is for OK
                if (result == VSConstants.MessageBoxResult.IDOK)
                {
                    try
                    {
                        // Delete file from disk
                        // Must use Directory.Delete to delete directory and content
                        if (macro.IsDirectory)
                        {
                            Directory.Delete(path, true);
                        }
                        else
                        {
                            File.Delete(path);
                        }

                        // Delete file from collection
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

        public void NewMacro()
        {
            MacroFSNode macro = this.SelectedMacro;
            macro.IsExpanded = true;

            string basePath = Path.Combine(macro.FullPath, "New Macro");
            string extension = ".js";

            string path = basePath + extension;

            // Increase count until filename is available (e.g. 'New Macro (2).js')
            int count = 2;
            while (File.Exists(path))
            {
                path = basePath + " (" + count++ + ")" + extension;
            }

            // Create the file
            File.WriteAllText(path, "/// <reference path=\"" + Manager.dteIntellisensePath + "\" />");

            // Refresh the tree
            this.Refresh();

            // Select new node
            MacroFSNode node = MacroFSNode.FindNodeFromFullPath(path);
            node.IsSelected = true;
            node.IsExpanded = true;
            node.IsEditable = true;
        }

        public void NewFolder()
        {
            MacroFSNode macro = this.SelectedMacro;
            macro.IsExpanded = true;

            string basePath = Path.Combine(macro.FullPath, "New Folder");
            string path = basePath;
            
            int count = 2;
            while (Directory.Exists(path))
            {
                path = path.Substring(0, path.Length - 4) + " (" + count++ + ")";
            }

            Directory.CreateDirectory(path);
            this.Refresh();

            MacroFSNode node = MacroFSNode.FindNodeFromFullPath(path);
            node.IsSelected = true;
            node.IsEditable = true;
        }

        public void CreateFileSystem()
        {
            // Create macro directory
            if (!Directory.Exists(VSMacrosPackage.Current.MacroDirectory))
            {
                Directory.CreateDirectory(VSMacrosPackage.Current.MacroDirectory);
            }

            // Create current macro file
            if (!File.Exists(Manager.CurrentMacroPath))
            {
                File.WriteAllText(Manager.CurrentMacroPath, "/// <reference path=\"" + Manager.dteIntellisensePath + "\" />");
            }

            // Create shortcuts file
            this.CreateShortcutFile();
        }

        public void Close()
        {
            this.SaveShortcuts();
        }

        public void MoveItem(MacroFSNode sourceItem, MacroFSNode targetItem)
        {
            string sourcePath = sourceItem.FullPath;
            string targetPath = Path.Combine(targetItem.FullPath, sourceItem.Name);
            string extension = ".js";

            MacroFSNode selected;

            // We want to expand the node and all its parents if it was expanded before OR if it is a file
            bool wasExpanded = sourceItem.IsExpanded || !sourceItem.IsDirectory;

            try
            {
                // Move on disk
                if (sourceItem.IsDirectory)
                {
                    System.IO.Directory.Move(sourcePath, targetPath);
                }
                else
                {
                    targetPath = targetPath + extension;
                    System.IO.File.Move(sourcePath, targetPath);
                }

                // Move shortcut as well
                if (sourceItem.Shortcut != MacroFSNode.NONE)
                {
                    int shortcutNumber = sourceItem.Shortcut;
                    Manager.Instance.Shortcuts[shortcutNumber] = targetPath;
                }                
            }
            catch (Exception e)
            {
                targetPath = sourceItem.FullPath;

                Manager.Instance.ShowMessageBox(e.Message);
            }
            finally
            {
                // Refresh tree
                Manager.Instance.Refresh();

                // Restore previously selected node
                selected = MacroFSNode.FindNodeFromFullPath(targetPath);
                selected.IsSelected = true;
                selected.IsExpanded = wasExpanded;

                // Notify change in shortcut
                selected.Shortcut = MacroFSNode.TO_FETCH;
            }

            // Make editable if the macro is the current macro
            if (sourceItem.FullPath == Manager.CurrentMacroPath)
            {
                selected.IsEditable = true;
            }

        }
        #region Helper Methods

        private StreamReader LoadFile(string path)
        {
            try
            {
                if (!File.Exists(path))
                {
                    throw new FileNotFoundException(Resources.MacroNotFound);
                }

                StreamReader str = new StreamReader(path);

                return str;
            }
            catch (FileNotFoundException e)
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
            this.shortcutsLoaded = true;

            try
            {
                // Get the path to the shortcut file
                string path = Path.Combine(VSMacrosPackage.Current.MacroDirectory, this.shortcutsFilePath);

                // If the file doesn't exist, initialize the Shortcuts array with empty strings
                if (!File.Exists(path))
                {
                    this.Shortcuts = Enumerable.Repeat(string.Empty, 10).ToArray();
                }
                else
                {
                    // Otherwise, load it
                    // Load XML file
                    var root = XDocument.Load(path);

                    // Parse to dictionary
                    this.Shortcuts = root.Descendants("command")
                                          .Select(elmt => elmt.Value)
                                          .ToArray();
                }
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

            xmlShortcuts.Save(this.shortcutsFilePath);
        }

        private void CreateShortcutFile()
        {
            string shortcutsPath = Path.Combine(VSMacrosPackage.Current.MacroDirectory, Manager.ShortcutsFileName);
            if (!File.Exists(shortcutsPath))
            {
                // Create file for writing UTF-8 encoded text
                File.WriteAllText(shortcutsPath, "<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\"?><commands><command>Command not bound. Do not use.</command><command/><command/><command/><command/><command/><command/><command/><command/><command/></commands>");
            }
        }

        public VSConstants.MessageBoxResult ShowMessageBox(string message, OLEMSGBUTTON btn = OLEMSGBUTTON.OLEMSGBUTTON_OK)
        {
            if (!this.uiShellLoaded)
            {
                return VSConstants.MessageBoxResult.IDABORT;
            }

            Guid clsid = Guid.Empty;
            int result;

            Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(
              this.uiShell.ShowMessageBox(
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

            return (VSConstants.MessageBoxResult)result;
        }
        #endregion
    }
}
