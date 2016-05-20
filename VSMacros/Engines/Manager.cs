//-----------------------------------------------------------------------
// <copyright file="Manager.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Threading;
using System.Xml.Linq;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using VSMacros.Dialogs;
using VSMacros.Interfaces;
using VSMacros.Models;
using Task = System.Threading.Tasks.Task;

namespace VSMacros.Engines
{
    public sealed class Manager
    {
        private static Manager instance;
        internal Executor executor;
        private Executor Executor
        {
            get
            {
                return (executor ?? (executor = new Executor()));
            }
        }

        #region Paths

        private const string CurrentMacroFileName = "Current.js";
        private const string ShortcutsFileName = "Shortcuts.xml";
        public const string IntellisenseFileName = "dte.js";
        private const string FolderExpansionFileName = "FolderExpansion.xml";

        public static string MacrosPath
        {
            get
            {
                return Path.Combine(VSMacrosPackage.Current.MacroDirectory, "Macros");
            }
        }

        public static string CurrentMacroPath {
            get 
            { 
                return Path.Combine(Manager.MacrosPath, Manager.CurrentMacroFileName);
            } 
        }
        
        public static string SamplesFolderPath 
        { 
            get 
            { 
                return Path.Combine(Manager.MacrosPath, "Samples"); 
            } 
        }

        public static string IntellisensePath
        {
            get
            {
                return Path.Combine(VSMacrosPackage.Current.MacroDirectory, Manager.IntellisenseFileName); 
            }
        }

        public static string ShortcutsPath
        {
            get
            {
                return Path.Combine(VSMacrosPackage.Current.MacroDirectory, Manager.ShortcutsFileName); 
            }
        }

        #endregion


        public static string[] Shortcuts { get; private set; }
        private bool shortcutsLoaded;
        private bool shortcutsDirty;

        private IServiceProvider serviceProvider;
        private IVsUIShell uiShell;
        private EnvDTE.DTE dte;

        private IRecorder recorder;

        private MacroFSNode SelectedMacro
        {
            get { return MacrosControl.Current != null ? MacrosControl.Current.SelectedNode : null; }
        }

        private Manager(IServiceProvider provider)
        {
            this.serviceProvider = provider;
            this.uiShell = (IVsUIShell)provider.GetService(typeof(SVsUIShell));
            this.dte = (EnvDTE.DTE)provider.GetService(typeof(SDTE));
            this.recorder = (IRecorder)this.serviceProvider.GetService(typeof(IRecorder));

            this.LoadShortcuts();
            this.shortcutsLoaded = true;
            this.shortcutsDirty = false;

            CreateFileSystem();
        }
        
        private static void AttachEvents(Executor executor)
        {
            executor.ResetMessages();

            var dispatcher = Dispatcher.CurrentDispatcher;

            executor.Complete += (sender, eventInfo) =>
            {
                if (eventInfo.IsError)
                {
                    Manager.Instance.ShowMessageBox(eventInfo.ErrorMessage);
                }

                Manager.instance.Executor.IsEngineRunning = false;

                ResetToolbar(dispatcher);
            };
        }

        private static void ResetToolbar(Dispatcher dispatcher)
        {
            try
            {
                dispatcher.Invoke(new Action(() =>
                {
                    VSMacrosPackage.Current.ClearStatusBar();
                    VSMacrosPackage.Current.UpdateButtonsForPlayback(false);
                }));
            }
            catch (Exception)
            {
                // Visual Studio is closing during execution.
            }
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

        public IVsWindowFrame PreviousWindow { get; set; }
        public bool PreviousWindowIsDocument { get; set; }

        public bool IsRecording { get; private set; }

        public void StartRecording()
        {
            // Move focus back to previous window
            if (this.dte.ActiveWindow.Caption == "Macro Explorer" && PreviousWindow != null)
            {
                PreviousWindow.Show();
            }

            this.PreviousWindowIsDocument = this.dte.ActiveWindow.Kind == "Document";

            this.IsRecording = true;
            this.recorder.StartRecording();
        }


        public void StopRecording()
        {
            string current = Manager.CurrentMacroPath;

            bool currentWasOpen = false;

            // Close current macro if open
            try
            {
                this.dte.Documents.Item(Manager.CurrentMacroPath).Close(EnvDTE.vsSaveChanges.vsSaveChangesNo);
                currentWasOpen = true;
            }
            catch (Exception e)
            {
                if (ErrorHandler.IsCriticalException(e))
                {
                    throw;
                }
            }

            this.recorder.StopRecording(current);
            this.IsRecording = false;

            MacroFSNode.SelectNode(CurrentMacroPath);

            // Reopen current macro
            if (currentWasOpen)
            {
                VsShellUtilities.OpenDocument(VSMacrosPackage.Current, Manager.CurrentMacroPath);
                this.PreviousWindow.Show();
            }
        }

        public void Playback(string path, int iterations = 1)
        {
            if (string.IsNullOrEmpty(path))
            {
                path = SelectedMacro == null ? CurrentMacroPath : SelectedMacro.FullPath;
            }

            // Before playing back, save the macro file
            this.SaveMacroIfDirty(path);

            // Move focus to first window
            if (this.dte.ActiveWindow.Caption == "Macro Explorer" && PreviousWindow != null)
            {
                PreviousWindow.Show();
            }

            VSMacrosPackage.Current.StatusBarChange(Resources.StatusBarPlayingText, 1);

            this.TogglePlayback(path, iterations);
        }

        public void PlaybackMultipleTimes(string path)
        {
            PlaybackMultipleTimesDialog dlg = new PlaybackMultipleTimesDialog();
            bool? result = dlg.ShowDialog();

            if (result.HasValue && result.Value)
            {
                int iterations;
                if (int.TryParse(dlg.IterationsTextbox.Text, out iterations)) 
                {
                    this.Playback(string.Empty, dlg.Iterations);
                }
            }
        }

        private void TogglePlayback(string path, int iterations)
        {
            AttachEvents(this.Executor);

            if (Manager.Instance.executor.IsEngineRunning)
            {
                VSMacrosPackage.Current.ClearStatusBar();
                Manager.Instance.executor.StopEngine();
            }
            else
            {
                VSMacrosPackage.Current.UpdateButtonsForPlayback(true);
                this.Executor.RunEngine(iterations, path);
                Manager.instance.Executor.CurrentlyExecutingMacro = this.GetExecutingMacroNameForPossibleErrorDisplay(this.SelectedMacro, path);
            }
        }

        private string GetExecutingMacroNameForPossibleErrorDisplay(MacroFSNode node, string path)
        {
            if(node != null)
            {
                return node.Name;
            }
            
            if(path == null)
            {
                throw new ArgumentNullException("path");
            }

            int lastBackslash = path.LastIndexOf('\\');
            string fileName = Path.GetFileNameWithoutExtension(path.Substring(lastBackslash != -1 ? lastBackslash + 1 : 0));

            return fileName;
        }

        public void PlaybackCommand(int cmd)
        {
            // Load shortcuts if not already loaded
            if (!this.shortcutsLoaded)
            {
                this.LoadShortcuts();
            }

            // Get path to macro bound to the shortcut
            string path = Manager.Shortcuts[cmd];

            if (!string.IsNullOrEmpty(path))
            {
                this.Playback(path, 1);
            }
        }

        public void StopPlayback()
        {
        }

        public void OpenFolder(string path = null)
        {
            path = !string.IsNullOrEmpty(path) ? path : this.SelectedMacro.FullPath;

            // Open the macro directory and let the user manage the macros
            System.Threading.Tasks.Task.Run(() => System.Diagnostics.Process.Start(path));
        }

        public void SaveCurrent()
        {
            SaveCurrentDialog dlg = new SaveCurrentDialog();
            dlg.ShowDialog();

            if (dlg.DialogResult == true)
            {
                try
                {
                    string pathToNew = Path.Combine(Manager.MacrosPath, dlg.MacroName.Text + ".js");
                    string pathToCurrent = Manager.CurrentMacroPath;

                    int newShortcutNumber = dlg.SelectedShortcutNumber;

                    // Move Current to new file and create a new Current
                    File.Move(pathToCurrent, pathToNew);
                    CreateCurrentMacro();

                    MacroFSNode macro = new MacroFSNode(pathToNew, MacroFSNode.RootNode);

                    if (newShortcutNumber != MacroFSNode.None)
                    {
                        // Update dictionary
                        Manager.Shortcuts[newShortcutNumber] = macro.FullPath;
                    }

                    this.SaveShortcuts(true);

                    this.Refresh();

                    // Select new node
                    MacroFSNode.SelectNode(pathToNew);
                }
                catch (Exception e)
                {
                    if (ErrorHandler.IsCriticalException(e))
                    {
                        throw;
                    }

                    this.ShowMessageBox(e.Message);
                }
            }
        }

        public void Refresh(bool reloadShortcut = true)
        {
            // If the shortcuts have been modified, ask to save them
            if (this.shortcutsDirty && reloadShortcut)
            {
                VSConstants.MessageBoxResult result = this.ShowMessageBox(Resources.ShortcutsChanged, OLEMSGBUTTON.OLEMSGBUTTON_YESNOCANCEL);

                switch (result)
                {
                    case VSConstants.MessageBoxResult.IDCANCEL:
                        return;
                    case VSConstants.MessageBoxResult.IDYES:
                        this.SaveShortcuts();
                        break;
                }
            }

            // Recreate file system to ensure that the required files exist
            CreateFileSystem();

            MacroFSNode.RefreshTree();

            this.LoadShortcuts();
        }

        public void Edit()
        {
            // TODO detect when a macro is dragged and it's opened -> use the overload to get the itemID
            MacroFSNode macro = this.SelectedMacro;
            string path = macro.FullPath;

            VsShellUtilities.OpenDocument(VSMacrosPackage.Current, path);
        }

        public void Rename()
        {
            MacroFSNode macro = this.SelectedMacro;

            if (macro.FullPath != Manager.CurrentMacroPath)
            {
                macro.EnableEdit();
            }
        }

        public void AssignShortcut()
        {
            AssignShortcutDialog dlg = new AssignShortcutDialog();
            dlg.ShowDialog();

            if (dlg.DialogResult == true)
            {
                MacroFSNode macro = this.SelectedMacro;

                // Remove old shortcut if it exists
                if (macro.Shortcut != MacroFSNode.None)
                {
                    Manager.Shortcuts[macro.Shortcut] = string.Empty;
                }

                int newShortcutNumber = dlg.SelectedShortcutNumber;

                // At this point, the shortcut has been removed
                // Assign a new one only if the user selected a key binding
                if (newShortcutNumber != MacroFSNode.None)
                {
                    // Get the node that previously owned that shortcut
                    MacroFSNode previousNode = MacroFSNode.FindNodeFromFullPath(Manager.Shortcuts[newShortcutNumber]);

                    // Update dictionary
                    Manager.Shortcuts[newShortcutNumber] = macro.FullPath;

                    // Update the UI binding for the old node
                    previousNode.Shortcut = MacroFSNode.ToFetch;
                }

                // Update UI with new macro's shortcut
                macro.Shortcut = MacroFSNode.ToFetch;

                // Mark the shortcuts in memory as dirty
                this.shortcutsDirty = true;
            }
        }

        public void SetShortcutsDirty()
        {
            this.shortcutsDirty = true;
        }

        public void Delete()
        {
            MacroFSNode macro = this.SelectedMacro;

            // Don't delete if macro is being edited
            if (macro.IsEditable)
            {
                return;
            }

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

                if (result == VSConstants.MessageBoxResult.IDOK)
                {
                    try
                    {
                        // Delete file or directory from disk
                        Manager.DeleteFileOrFolder(path);

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
            File.WriteAllText(path, "/// <reference path=\"" + Manager.IntellisensePath + "\" />");

            // Refresh the tree
            MacroFSNode.RefreshTree();

            // Select new node
            MacroFSNode node = MacroFSNode.SelectNode(path);
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
                path = basePath + " (" + count++ + ")";
            }

            Directory.CreateDirectory(path);
            MacroFSNode.RefreshTree();

            MacroFSNode node = MacroFSNode.SelectNode(path);
            node.IsEditable = true;
        }

        public static void CreateFileSystem()
        {
            // Create main macro directory
            if (!Directory.Exists(VSMacrosPackage.Current.MacroDirectory))
            {
                Directory.CreateDirectory(VSMacrosPackage.Current.MacroDirectory);
            }

            // Create macros folder directory
            if (!Directory.Exists(Manager.MacrosPath))
            {
                Directory.CreateDirectory(Manager.MacrosPath);
            }

            // Create current macro file
            CreateCurrentMacro();

            // Create shortcuts file
            CreateShortcutFile();

            // Copy Samples folder
            string samplesTargetDir = SamplesFolderPath;
            if (!Directory.Exists(samplesTargetDir))
            {
                string samplesSourceDir = Path.Combine(VSMacrosPackage.Current.AssemblyDirectory, "Macros", "Samples");
                DirectoryCopy(samplesSourceDir, samplesTargetDir, true);
            }

            // Copy DTE IntelliSense file
            string dteFileTargetPath = IntellisensePath;
            if (!File.Exists(dteFileTargetPath))
            {
                string dteFileSourcePath = Path.Combine(VSMacrosPackage.Current.AssemblyDirectory, "Intellisense", IntellisenseFileName);
                File.Copy(dteFileSourcePath, dteFileTargetPath);
            }
        }

        public void Close()
        {
            this.SaveFolderExpansion();
            this.SaveShortcuts();
        }

        private string RelativeIntellisensePath(int depth)
        {
            string path = Manager.IntellisenseFileName;

            for (int i = 0; i < depth; i++)
            {
                path = "../" + path;
            }

            return path;
        }

        public void MoveItem(MacroFSNode sourceItem, MacroFSNode targetItem)
        {
            string sourcePath = sourceItem.FullPath;
            string targetPath = Path.Combine(targetItem.FullPath, sourceItem.Name);
            string extension = ".js";

            MacroFSNode selected;

            // We want to expand the node and all its parents if it was expanded before OR if it is a file
            bool wasExpanded = sourceItem.IsExpanded;

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

                    // Close in the editor
                    this.Reopen(sourcePath, targetPath);
                }

                // Move shortcut as well
                if (sourceItem.Shortcut != MacroFSNode.None)
                {
                    int shortcutNumber = sourceItem.Shortcut;
                    Manager.Shortcuts[shortcutNumber] = targetPath;
                }
            }
            catch (Exception e)
            {
                if (ErrorHandler.IsCriticalException(e))
                {
                    throw;
                }

                targetPath = sourceItem.FullPath;

                Manager.Instance.ShowMessageBox(e.Message);
            }

            CreateCurrentMacro();

            // Refresh tree
            MacroFSNode.RefreshTree();

            // Restore previously selected node
            selected = MacroFSNode.SelectNode(targetPath);
            selected.IsExpanded = wasExpanded;
            selected.Parent.IsExpanded = true;

            // Notify change in shortcut
            selected.Shortcut = MacroFSNode.ToFetch;

            // Make editable if the macro is the current macro
            if (sourceItem.FullPath == Manager.CurrentMacroPath)
            {
                selected.IsEditable = true;
            }
        }

        #region Helper Methods

        public void Reopen(string source, string target)
        {
            try
            {
                this.dte.Documents.Item(source).Close(EnvDTE.vsSaveChanges.vsSaveChangesNo);
                this.dte.ItemOperations.OpenFile(target);
            }
            catch (ArgumentException) { }
        }

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

        public void LoadFolderExpansion()
        {
            string path = Path.Combine(VSMacrosPackage.Current.UserLocalDataPath, Manager.FolderExpansionFileName);

            if (File.Exists(path))
            {
                HashSet<string> enabledDirs = new HashSet<string>();

                string[] folders = File.ReadAllText(path).Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

                foreach (var s in folders)
                {
                    enabledDirs.Add(s);
                }

                MacroFSNode.EnabledDirectories = enabledDirs;
            }
        }

        private void SaveFolderExpansion()
        {
            var folders = string.Join(Environment.NewLine, MacroFSNode.EnabledDirectories);

            File.WriteAllText(Path.Combine(VSMacrosPackage.Current.UserLocalDataPath, Manager.FolderExpansionFileName), folders);
        }

        private void LoadShortcuts()
        {
            this.shortcutsLoaded = true;

            try
            {
                // Get the path to the shortcut file
                string path = Manager.ShortcutsPath;

                // If the file doesn't exist, initialize the Shortcuts array with empty strings
                if (!File.Exists(path))
                {
                    Manager.Shortcuts = Enumerable.Repeat(string.Empty, 10).ToArray();
                }
                else
                {
                    // Otherwise, load it
                    // Load XML file
                    var root = XDocument.Load(path);

                    // Parse to dictionary
                    Manager.Shortcuts = root.Descendants("command")
                                          .Select(elmt => elmt.Value)
                                          .ToArray();
                }
            }
            catch (Exception e)
            {
                this.ShowMessageBox(e.Message);
            }
        }

        public void SaveShortcuts(bool overwrite = false)
        {
            if (this.shortcutsDirty || overwrite)
            {
                XDocument xmlShortcuts =
                    new XDocument(
                        new XDeclaration("1.0", "utf-8", "yes"),
                        new XElement("commands",
                            from s in Manager.Shortcuts
                            select new XElement("command",
                                new XText(s))));

                xmlShortcuts.Save(Manager.ShortcutsPath);

                this.shortcutsDirty = false;
            }
        }

        private static void CreateCurrentMacro()
        {
            if (!File.Exists(Manager.CurrentMacroPath))
            {
                File.Create(Manager.CurrentMacroPath).Close();

                Task.Run(() =>
                {
                    // Write to current macro file
                    File.WriteAllText(Manager.CurrentMacroPath, "/// <reference path=\"" + Manager.IntellisensePath + "\" />");
                });
            }
        }

        private static void CreateShortcutFile()
        {
            string shortcutsPath = Manager.ShortcutsPath;
            if (!File.Exists(shortcutsPath))
            {
                // Create file for writing UTF-8 encoded text
                File.WriteAllText(shortcutsPath, "<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\"?><commands><command>Command not bound. Do not use.</command><command/><command/><command/><command/><command/><command/><command/><command/><command/></commands>");
            }
        }

        private void SaveMacroIfDirty(string path)
        {
            try
            {
                EnvDTE.Document doc = this.dte.Documents.Item(path);

                if (!doc.Saved)
                {
                    if (VSConstants.MessageBoxResult.IDYES == this.ShowMessageBox(
                        string.Format(Resources.MacroNotSavedBeforePlayback, Path.GetFileNameWithoutExtension(path)),
                        OLEMSGBUTTON.OLEMSGBUTTON_YESNO))
                    {
                        doc.Save();
                    }
                }
            }
            catch (Exception e)
            {
                if (ErrorHandler.IsCriticalException(e)) { throw; }
            }
        }

        public VSConstants.MessageBoxResult ShowMessageBox(string message, OLEMSGBUTTON btn = OLEMSGBUTTON.OLEMSGBUTTON_OK)
        {
            if (this.uiShell == null)
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

        #region Disk Operations

        private const int FO_DELETE = 0x0003;
        private const int FOF_ALLOWUNDO = 0x0040;           // Preserve undo information, if possible. 
        private const int FOF_NOCONFIRMATION = 0x0010;      // Show no confirmation dialog box to the user

        // Struct which contains information that the SHFileOperation function uses to perform file operations. 
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct SHFILEOPSTRUCT
        {
            public IntPtr hwnd;
            [MarshalAs(UnmanagedType.U4)]
            public int wFunc;
            public string pFrom;
            public string pTo;
            public short fFlags;
            [MarshalAs(UnmanagedType.Bool)]
            public bool fAnyOperationsAborted;
            public IntPtr hNameMappings;
            public string lpszProgressTitle;
        }

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        static extern int SHFileOperation(ref SHFILEOPSTRUCT FileOp);

        public static void DeleteFileOrFolder(string path)
        {
            SHFILEOPSTRUCT fileop = new SHFILEOPSTRUCT();
            fileop.wFunc = FO_DELETE;
            fileop.pFrom = path + '\0' + '\0';
            fileop.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION;
            SHFileOperation(ref fileop);
        }

        public static void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs)
        {
            // Get the subdirectories for the specified directory.
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);
            DirectoryInfo[] dirs = dir.GetDirectories();

            if (!dir.Exists)
            {
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found: "
                    + sourceDirName);
            }

            // If the destination directory doesn't exist, create it. 
            if (!Directory.Exists(destDirName))
            {
                Directory.CreateDirectory(destDirName);
            }

            // Get the files in the directory and copy them to the new location.
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string temppath = Path.Combine(destDirName, file.Name);

                if (!File.Exists(temppath))
                {
                    file.CopyTo(temppath, false);
                }
            }

            // If copying subdirectories, copy them and their contents to new location. 
            if (copySubDirs)
            {
                foreach (DirectoryInfo subdir in dirs)
                {
                    string temppath = Path.Combine(destDirName, subdir.Name);
                    DirectoryCopy(subdir.FullName, temppath, copySubDirs);
                }
            }
        }

        #endregion
    }
}
