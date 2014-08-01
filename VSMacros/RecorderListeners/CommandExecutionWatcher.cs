//-----------------------------------------------------------------------
// <copyright file="CommandExecutionWatcher.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using Microsoft.Internal.VisualStudio.Shell;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell.Interop;
using VSMacros.Interfaces;
using IServiceProvider = System.IServiceProvider;
using OLEConstants = Microsoft.VisualStudio.OLE.Interop.Constants;

namespace VSMacros.RecorderListeners
{
    // This class will hook into the command route and be responsible for monitoring commands executions.
    internal sealed class CommandExecutionWatcher : IOleCommandTarget, IDisposable
    {
        private const string UnknownCommand = null;

        private IServiceProvider serviceProvider;
        private uint priorityCommandTargetCookie;
        private IRecorderPrivate macroRecorder;

        internal CommandExecutionWatcher(IServiceProvider serviceProvider)
        {
            Validate.IsNotNull(serviceProvider, "serviceProvider");
            this.serviceProvider = serviceProvider;
         
            var rpct = (IVsRegisterPriorityCommandTarget)this.serviceProvider.GetService(typeof(SVsRegisterPriorityCommandTarget));
            if (rpct != null)
            {
                // We can ignore the return code here as there really isn't anything reasonable we could do to deal with failure, 
                // and it is essentially a no-fail method.
                rpct.RegisterPriorityCommandTarget(dwReserved: 0U, pCmdTrgt: this, pdwCookie: out this.priorityCommandTargetCookie);
            }
            this.macroRecorder = (IRecorderPrivate)serviceProvider.GetService(typeof(IRecorder));
        }

        public int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            if (this.macroRecorder.IsRecording)
            {
                // An Exec call with a non-null pvaOut implies it is actually the shell trying to get the combo box child items for a 
                // combo, not a real command execution, so we can ignore these for purposes of command recording.
                if (pvaOut == IntPtr.Zero && (pguidCmdGroup != GuidList.GuidVSMacrosCmdSet || nCmdID != PkgCmdIDList.CmdIdRecord) &&
                    !(pguidCmdGroup == new Guid("{5efc7975-14bc-11cf-9b2b-00aa00573819}") && nCmdID == 770))
                {
                    string commandName = this.ConvertGuidDWordToName(pguidCmdGroup, nCmdID);
                    this.macroRecorder.AddCommandData(pguidCmdGroup, nCmdID, commandName, (char)0);
                }
            }
            
            // We never actually handle Exec (i.e. return S_OK) because we don't want to claim we have handled the execution of any commands, 
            // we are just watching them go by.
            return (int)OLEConstants.OLECMDERR_E_NOTSUPPORTED;
        }

        public int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
        {
            // We never handle query status, we don't want to affect the enabled/visible state of any commands, just watch execution requests.
            return (int)OLEConstants.OLECMDERR_E_NOTSUPPORTED;
        }

        public void Dispose()
        {
            if (this.priorityCommandTargetCookie != 0U && this.serviceProvider != null)
            {
                var rpct = (IVsRegisterPriorityCommandTarget)this.serviceProvider.GetService(typeof(SVsRegisterPriorityCommandTarget));
                if (rpct != null)
                {
                    // We can ignore the return code here as there really isn't anything reasonable we could do to deal with failure, 
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
