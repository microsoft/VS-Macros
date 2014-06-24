using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.IO;
using VSMacros.Models;

namespace VSMacros.Engines
{
    interface IManager
    {
        void ToggleRecording();

        void Playback(string path, int times);

        void StopPlayback();

        void SaveCurrent();

        void Edit();

        void Rename();

        void AssignShortcut();

        void Delete();
    }

    public sealed class Manager : IManager
    {
        private static readonly Manager instance = new Manager();

        private Manager() 
        { 
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
        }

        public void StopPlayback() 
        {
        }

        public void SaveCurrent() 
        {
        }

        public void Edit()
        {
            MacroFSNode macro = SelectedMacro;
            string path = macro.FullPath;

            EnvDTE.DTE dte = (EnvDTE.DTE)System.Runtime.InteropServices.Marshal.GetActiveObject("VisualStudio.DTE.12.0");
            dte.ItemOperations.OpenFile(path);
        }

        public void Rename()
        {
            MacroFSNode macro = SelectedMacro;
            macro.EnableEdit();
        }

        public void AssignShortcut()
        {
        }

        public void Delete()
        {
            MacroFSNode macro = SelectedMacro;
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
                IVsUIShell uiShell = (IVsUIShell)((IServiceProvider)VSMacrosPackage.Current).GetService(typeof(SVsUIShell));
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
                    OLEMSGBUTTON.OLEMSGBUTTON_OKCANCEL,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST,
                    OLEMSGICON.OLEMSGICON_WARNING,
                    0,        // false
                    out result));

                // TODO 1 is for OK (replace 1 by IDOK)
                if (result == 1)
                {
                    file.Delete();  // TODO non-empty dir will raise an exception here -> User Directory.Delete(path, true)
                    macro.Delete();
                }
            }
            else
            {
                macro.Delete();
            }
        }
      
        private FileStream LoadFile(string path) 
        { 
            return null; 
        }

        private void SaveToFile(Stream str, string path)
        { 
        }

        private MacroFSNode SelectedMacro
        {
            get { return MacrosControl.Current.SelectedNode; }
        }
    }
}
