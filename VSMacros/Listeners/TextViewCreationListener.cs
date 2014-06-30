using System;
using System.ComponentModel.Composition;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Utilities;
using System.Windows.Forms;
using Microsoft.VisualStudio.Shell.Interop;
using IServiceProvider = System.IServiceProvider;
using Microsoft.Internal.VisualStudio.Shell;

using OLEConstants = Microsoft.VisualStudio.OLE.Interop.Constants;
using Microsoft.VisualStudio.Shell;

namespace VSMacros
{
    internal class EditorCommandFilter : IOleCommandTarget
    {
        private SVsServiceProvider serviceProvider;
        private IMacroRecorderPrivate macroRecorder;

        internal IOleCommandTarget NextCommandTarget { get; set; }

        internal EditorCommandFilter(SVsServiceProvider serviceProvider)
        {
            this.serviceProvider = serviceProvider;
            macroRecorder = (IMacroRecorderPrivate)serviceProvider.GetService(typeof(IMacroRecorder));
        }

        public int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            if (!macroRecorder.Recording)
            {   
                if (NextCommandTarget != null)
                {
                    return NextCommandTarget.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
                }
                return (int)OLEConstants.OLECMDERR_E_NOTSUPPORTED;
            }
            if ((pguidCmdGroup == VSConstants.CMDSETID.StandardCommandSet2K_guid) && (nCmdID == (uint)VSConstants.VSStd2KCmdID.TYPECHAR))
            {
                macroRecorder.AddCommandData(pguidCmdGroup, nCmdID, "Keyboard", (char)(ushort)Marshal.GetObjectForNativeVariant(pvaIn));
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
            ErrorHandler.ThrowOnFailure(textViewAdapter.AddCommandFilter(commandFilter, out nextTarget));
            commandFilter.NextCommandTarget = nextTarget;
        }
    }
}
