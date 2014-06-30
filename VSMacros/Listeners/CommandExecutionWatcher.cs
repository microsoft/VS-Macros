using System;
using Microsoft.Internal.VisualStudio.Shell;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell.Interop;
using IServiceProvider = System.IServiceProvider;
using OLEConstants = Microsoft.VisualStudio.OLE.Interop.Constants;

namespace VSMacros
{
    // NOTE: This class will hook into the command route and be responsible for monitoring commands executions.
    internal sealed class CommandExecutionWatcher : IOleCommandTarget, IDisposable
    {
        private const string UnknownCommand = "<Unknown>";

        private IServiceProvider serviceProvider;
        private uint priorityCommandTargetCookie;
        private IMacroRecorderPrivate macroRecorder;

        internal CommandExecutionWatcher(IServiceProvider serviceProvider)
        {
            Validate.IsNotNull(serviceProvider, "serviceProvider");
            this.serviceProvider = serviceProvider;
         
            var rpct = (IVsRegisterPriorityCommandTarget)this.serviceProvider.GetService(typeof(SVsRegisterPriorityCommandTarget));
            if (rpct != null)
            {
                // NOTE: We can ignore the return code here as there really isn't anything reasonable we could do to deal with failure, 
                // and it is essentially a no-fail method.
                rpct.RegisterPriorityCommandTarget(dwReserved: 0U, pCmdTrgt: this, pdwCookie: out this.priorityCommandTargetCookie);
            }
            macroRecorder = (IMacroRecorderPrivate)serviceProvider.GetService(typeof(IMacroRecorder));
        }

        public int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            if (macroRecorder.Recording)
            {
                // NOTE: An Exec call with a non-null pvaOut implies it is actually the shell trying to get the combo box child items for a 
                // combo, not a real command execution, so we can ignore these for purposes of command recording.
                if (pvaOut == IntPtr.Zero && (pguidCmdGroup != GuidList.guidVSMacrosCmdSet || nCmdID != PkgCmdIDList.cmdidRecord))
                {
                    string commandName = ConvertGuidDWordToName(pguidCmdGroup, nCmdID);
                    macroRecorder.AddCommandData(pguidCmdGroup, nCmdID, commandName, (char)0);
                }
            }
            // NOTE: We never actually handle Exec (i.e. return S_OK) because we don't want to claim we have handled the execution of any commands, 
            // we are just watching them go by.
            return (int)OLEConstants.OLECMDERR_E_NOTSUPPORTED;
        }

        public int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
        {
            // NOTE: We never handle query status, we don't want to affect the enabled/visible state of any commands, just watch execution requests.
            return (int)OLEConstants.OLECMDERR_E_NOTSUPPORTED;
        }

        public void Dispose()
        {
            if (this.priorityCommandTargetCookie != 0U && this.serviceProvider != null)
            {
                var rpct = (IVsRegisterPriorityCommandTarget)this.serviceProvider.GetService(typeof(SVsRegisterPriorityCommandTarget));
                if (rpct != null)
                {
                    // NOTE: We can ignore the return code here as there really isn't anything reasonable we could do to deal with failure, 
                    // and it is essentially a no-fail method.
                    rpct.UnregisterPriorityCommandTarget(this.priorityCommandTargetCookie);
                    this.priorityCommandTargetCookie = 0U;
                }

                this.serviceProvider = null;
            }
        }

        private string ConvertGuidDWordToName(Guid guid, uint dword)
        {
            var cmdNameMapping = (IVsCmdNameMapping)this.serviceProvider.GetService(typeof(SVsCmdNameMapping));
            if (cmdNameMapping == null)
            {
                return UnknownCommand;
            }

            string name;
            if (ErrorHandler.Failed(cmdNameMapping.MapGUIDIDToName(ref guid, dword, VSCMDNAMEOPTS.CNO_GETENU, out name)) ||
               string.IsNullOrEmpty(name))
            {
                return UnknownCommand;
            }

            return name;
        }
    }

}
