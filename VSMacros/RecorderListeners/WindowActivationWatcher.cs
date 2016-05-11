//-----------------------------------------------------------------------
// <copyright file="WindowActivationWatcher.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using Microsoft.Internal.VisualStudio.Shell;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using VSMacros.Engines;
using VSMacros.Interfaces;
using VSMacros.Model;

namespace VSMacros.RecorderListeners
{
    // NOTE: This class will hook into the selection events and be responsible for monitoring selection changes
    internal sealed class WindowActivationWatcher : IVsSelectionEvents, IDisposable
    {
        private IServiceProvider serviceProvider;
        private uint monSelCookie;
        private IRecorderPrivate macroRecorder;
        private RecorderDataModel dataModel;

        // NOTE: Values obtained from http://msdn.microsoft.com/en-us/library/microsoft.visualstudio.shell.interop.__vsfpropid.aspx, specifically the part for VSFPROPID_Type.
        enum FrameType
        {
            Document = 1,
            ToolWindow
        }

        internal WindowActivationWatcher(IServiceProvider serviceProvider, RecorderDataModel dataModel)
        {
            Validate.IsNotNull(serviceProvider, "serviceProvider");
            Validate.IsNotNull(dataModel, "dataModel");

            this.serviceProvider = serviceProvider;
            this.dataModel = dataModel;

            var monSel = (IVsMonitorSelection)this.serviceProvider.GetService(typeof(SVsShellMonitorSelection));
            if (monSel != null)
            {
                // NOTE: We can ignore the return code here as there really isn't anything reasonable we could do to deal with failure, 
                // and it is essentially a no-fail method.
                monSel.AdviseSelectionEvents(pSink: this, pdwCookie: out this.monSelCookie);
            }
            this.macroRecorder = (IRecorderPrivate)serviceProvider.GetService(typeof(IRecorder));
        }

        public int OnElementValueChanged(uint elementid, object varValueOld, object varValueNew)
        {
            var elementId = (VSConstants.VSSELELEMID)elementid;
            if (elementId == VSConstants.VSSELELEMID.SEID_WindowFrame)
            {
                if (varValueNew != null)
                {
                    // NOTE: We have a selection change to a non-null value, this means someone has switched the active document / toolwindow (or the shell has done
                    // so automatically since they closed the previously active one).
                    var windowFrame = (IVsWindowFrame)varValueNew;
                    var windowFrameOld = (IVsWindowFrame)varValueOld;
                    object untypedProperty;
                    object untypedPropertyOld;

                    if (ErrorHandler.Succeeded(windowFrame.GetProperty((int)__VSFPROPID.VSFPROPID_Type, out untypedProperty)))
                    {
                        FrameType typedProperty = (FrameType)(int)untypedProperty;

                        if (windowFrameOld != null)
                        {
                            if (ErrorHandler.Succeeded(windowFrameOld.GetProperty((int)__VSFPROPID.VSFPROPID_Caption, out untypedPropertyOld)))
                            {
                                string captionOld = (string)untypedPropertyOld;
                                if (captionOld != "Macro Explorer")
                                {
                                    Manager.Instance.PreviousWindow = windowFrameOld;
                                }
                            }

                            if (Manager.Instance.IsRecording)
                            {
                                if (typedProperty == FrameType.Document)
                                {
                                    if (ErrorHandler.Succeeded(windowFrame.GetProperty((int)__VSFPROPID.VSFPROPID_pszMkDocument, out untypedProperty)))
                                    {
                                        string docPath = (string)untypedProperty;
                                        if (!((this.dataModel.isDoc == true) && (this.dataModel.currDoc == docPath)))
                                        {
                                            this.dataModel.AddWindow(docPath);
                                        }
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
                                            if (caption != "Macro Explorer")
                                            {
                                                if (!((this.dataModel.isDoc == false) && (this.dataModel.currWindow == windowID)))
                                                {
                                                    this.dataModel.AddWindow(windowID, caption);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            return VSConstants.S_OK;
                        }
                    }
                }
            }
            return VSConstants.S_OK;
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
