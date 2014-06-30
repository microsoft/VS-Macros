using System;
using Microsoft.Internal.VisualStudio.Shell;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;

namespace VSMacros
{
    // NOTE: This class will hook into the selection events and be responsible for monitoring selection changes
    internal sealed class WindowActivationWatcher : IVsSelectionEvents, IDisposable
    {
        private IServiceProvider serviceProvider;
        private uint monSelCookie;
        private IMacroRecorderPrivate macroRecorder;

        // NOTE: Values obtained from http://msdn.microsoft.com/en-us/library/microsoft.visualstudio.shell.interop.__vsfpropid.aspx, specifically the part for VSFPROPID_Type.
        enum FrameType
        {
            Document = 1,
            ToolWindow
        }

        internal WindowActivationWatcher(IServiceProvider serviceProvider)
        {
            Validate.IsNotNull(serviceProvider, "serviceProvider");

            this.serviceProvider = serviceProvider;

            var monSel = (IVsMonitorSelection)this.serviceProvider.GetService(typeof(SVsShellMonitorSelection));
            if (monSel != null)
            {
                // NOTE: We can ignore the return code here as there really isn't anything reasonable we could do to deal with failure, 
                // and it is essentially a no-fail method.
                monSel.AdviseSelectionEvents(pSink: this, pdwCookie: out this.monSelCookie);
            }
            macroRecorder = (IMacroRecorderPrivate)serviceProvider.GetService(typeof(IMacroRecorder));
        }

        public int OnElementValueChanged(uint elementid, object varValueOld, object varValueNew)
        {
            if (!macroRecorder.Recording)
            {
                return VSConstants.S_OK;
            }

            var elementId = (VSConstants.VSSELELEMID)elementid;
            if (elementId == VSConstants.VSSELELEMID.SEID_WindowFrame)
            {
                if (varValueNew != null)
                {
                    // NOTE: We have a selection change to a non-null value, this means someone has switched the active document / toolwindow (or the shell has done
                    // so automatically since they closed the previously active one).
                    var windowFrame = (IVsWindowFrame)varValueNew;
                    object untypedProperty;
                    if (ErrorHandler.Succeeded(windowFrame.GetProperty((int)__VSFPROPID.VSFPROPID_Type, out untypedProperty)))
                    {
                        FrameType typedProperty = (FrameType)(int)untypedProperty;
                        if (typedProperty == FrameType.Document)
                        {
                            if (ErrorHandler.Succeeded(windowFrame.GetProperty((int)__VSFPROPID.VSFPROPID_pszMkDocument, out untypedProperty)))
                            {
                                string docPath = (string)untypedProperty;
                                this.macroRecorder.AddWindowActivation(docPath);
                            }
                        }
                        else if (typedProperty == FrameType.ToolWindow)
                        {
                            if (ErrorHandler.Succeeded(windowFrame.GetProperty((int)__VSFPROPID.VSFPROPID_Caption, out untypedProperty)))
                            {
                                string caption = (string)untypedProperty;
                                Guid windowID;
                                if (ErrorHandler.Succeeded(windowFrame.GetGuidProperty((int)__VSFPROPID.VSFPROPID_GuidPersistenceSlot, out windowID)))
                                {
                                    this.macroRecorder.AddWindowActivation(windowID, caption);
                                }

                            }
                        }
                    }

                }
            }
            return VSConstants.S_OK; ;
        }

        public int OnCmdUIContextChanged(uint dwCmdUICookie, int fActive)
        {
            // NOTE: We don't care about UI context changes like package loading, command visibility, etc.
            return VSConstants.S_OK;
        }

        public int OnSelectionChanged(IVsHierarchy pHierOld, uint itemidOld, IVsMultiItemSelect pMISOld, ISelectionContainer pSCOld, IVsHierarchy pHierNew, uint itemidNew, IVsMultiItemSelect pMISNew, ISelectionContainer pSCNew)
        {
            // NOTE: We don't care about selection changes like the solution explorer
            return VSConstants.S_OK;
        }

        public void Dispose()
        {
            if (this.monSelCookie != 0U && this.serviceProvider != null)
            {
                var monSel = (IVsMonitorSelection)this.serviceProvider.GetService(typeof(SVsShellMonitorSelection));
                if (monSel != null)
                {
                    // NOTE: We can ignore the return code here as there really isn't anything reasonable we could do to deal with failure, 
                    // and it is essentially a no-fail method.
                    monSel.UnadviseSelectionEvents(this.monSelCookie);
                    this.monSelCookie = 0U;
                }
            }

            this.serviceProvider = null;

        }
    }
}
