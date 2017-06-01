//-----------------------------------------------------------------------
// <copyright file="TextViewCreationListener.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.ComponentModel.Composition;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Utilities;

namespace VSMacros.RecorderListeners
{
    [Export(typeof(IVsTextViewCreationListener))]
    [TextViewRole(PredefinedTextViewRoles.Editable)]
    [ContentType("Text")]
    internal class TextViewCreationListener : IVsTextViewCreationListener
    {
        [Import]
        private SVsServiceProvider serviceProvider = null;

        public void VsTextViewCreated(IVsTextView textViewAdapter)
        {
            var commandFilter = new EditorCommandFilter(serviceProvider: this.serviceProvider);
            IOleCommandTarget nextTarget;
            ErrorHandler.ThrowOnFailure(textViewAdapter.AddCommandFilter(commandFilter, out nextTarget));
            commandFilter.NextCommandTarget = nextTarget;
        }
    }
}
