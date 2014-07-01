using System;
using System.ComponentModel.Composition;
using System.Runtime.InteropServices;
using Microsoft.Internal.VisualStudio.Shell;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Utilities;
using VSMacros.Interfaces;
using OLEConstants = Microsoft.VisualStudio.OLE.Interop.Constants;

namespace VSMacros
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
            if (macroRecorder.Recording)
            {
                if ((pguidCmdGroup == VSConstants.CMDSETID.StandardCommandSet2K_guid) && (nCmdID == (uint)VSConstants.VSStd2KCmdID.TYPECHAR))
                {
                    this.macroRecorder.AddCommandData(pguidCmdGroup, nCmdID, "Keyboard", (char)(ushort)Marshal.GetObjectForNativeVariant(pvaIn));
                }    
            }
            if (NextCommandTarget != null)
            {
                return NextCommandTarget.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
            }
            return (int)OLEConstants.OLECMDERR_E_NOTSUPPORTED;
        }

        public int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
        {
            if (NextCommandTarget != null)
            {
                return NextCommandTarget.QueryStatus(ref pguidCmdGroup, cCmds, prgCmds, pCmdText);
            }

            return (int)OLEConstants.OLECMDERR_E_NOTSUPPORTED;
        }
    }

    [Export(typeof(IVsTextViewCreationListener))]
    [TextViewRole(PredefinedTextViewRoles.Editable)]
    [ContentType("Text")]
    internal class TextViewCreationListener : IVsTextViewCreationListener
    {
        [Import]
        private SVsServiceProvider serviceProvider;
        private EditorCommandFilter commandFilter;

        public void VsTextViewCreated(IVsTextView textViewAdapter)
        {
            if (this.commandFilter == null)
            {
                this.commandFilter = new EditorCommandFilter(serviceProvider: this.serviceProvider);
            }
            IOleCommandTarget nextTarget;
            ErrorHandler.ThrowOnFailure(textViewAdapter.AddCommandFilter(this.commandFilter, out nextTarget));
            this.commandFilter.NextCommandTarget = nextTarget;
        }
    }
}
