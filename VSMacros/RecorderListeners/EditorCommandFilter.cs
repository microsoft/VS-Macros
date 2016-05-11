//-----------------------------------------------------------------------
// <copyright file="EditorCommandFilter.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Runtime.InteropServices;
using Microsoft.Internal.VisualStudio.Shell;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using VSMacros.Interfaces;
using OLEConstants = Microsoft.VisualStudio.OLE.Interop.Constants;

namespace VSMacros.RecorderListeners
{
    internal class EditorCommandFilter : IOleCommandTarget
    {
        private SVsServiceProvider serviceProvider;
        private IRecorderPrivate macroRecorder;

        internal IOleCommandTarget NextCommandTarget { get; set; }

        internal EditorCommandFilter(SVsServiceProvider serviceProvider)
        {
            Validate.IsNotNull(serviceProvider, "serviceProvider");
            this.serviceProvider = serviceProvider;
            this.macroRecorder = (IRecorderPrivate)serviceProvider.GetService(typeof(IRecorder));
        }

        public int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            if (macroRecorder.IsRecording)
            {
                if ((pguidCmdGroup == VSConstants.CMDSETID.StandardCommandSet2K_guid) && (nCmdID == (uint)VSConstants.VSStd2KCmdID.TYPECHAR))
                {
                    this.macroRecorder.AddCommandData(pguidCmdGroup, nCmdID, "keyboard", (char)(ushort)Marshal.GetObjectForNativeVariant(pvaIn));
                }
            }
            if (this.NextCommandTarget != null)
            {
                return this.NextCommandTarget.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
            }
            return (int)OLEConstants.OLECMDERR_E_NOTSUPPORTED;
        }

        public int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
        {
            if (this.NextCommandTarget != null)
            {
                return this.NextCommandTarget.QueryStatus(ref pguidCmdGroup, cCmds, prgCmds, pCmdText);
            }

            return (int)OLEConstants.OLECMDERR_E_NOTSUPPORTED;
        }
    }
}
