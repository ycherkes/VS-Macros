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
            // NOTE: We can ignore the return code here as there really isn't anything reasonable we could do to deal with failure,
            // and it is essentially a no-fail method.
            monSel?.AdviseSelectionEvents(pSink: this, pdwCookie: out monSelCookie);
            this.macroRecorder = (IRecorderPrivate)serviceProvider.GetService(typeof(IRecorder));
        }

        public int OnElementValueChanged(uint elementid, object varValueOld, object varValueNew)
        {
            var elementId = (VSConstants.VSSELELEMID)elementid;
            if (elementId != VSConstants.VSSELELEMID.SEID_WindowFrame) return VSConstants.S_OK;

            if (varValueNew == null) return VSConstants.S_OK;
            // NOTE: We have a selection change to a non-null value, this means someone has switched the active document / toolwindow (or the shell has done
            // so automatically since they closed the previously active one).
            var windowFrame = (IVsWindowFrame)varValueNew;
            var windowFrameOld = (IVsWindowFrame)varValueOld;
            object untypedProperty;

            if (!ErrorHandler.Succeeded(windowFrame.GetProperty((int) __VSFPROPID.VSFPROPID_Type, out untypedProperty)))
                return VSConstants.S_OK;

            var typedProperty = (FrameType)(int)untypedProperty;

            if (windowFrameOld != null)
            {
                object untypedPropertyOld;
                if (ErrorHandler.Succeeded(windowFrameOld.GetProperty((int)__VSFPROPID.VSFPROPID_Caption, out untypedPropertyOld)))
                {
                    var captionOld = (string)untypedPropertyOld;
                    if (captionOld != "Macro Explorer")
                    {
                        Manager.Instance.PreviousWindow = windowFrameOld;
                    }
                }

                if (!Manager.Instance.IsRecording) return VSConstants.S_OK;

                switch (typedProperty)
                {
                    case FrameType.Document:
                        if (ErrorHandler.Succeeded(windowFrame.GetProperty((int)__VSFPROPID.VSFPROPID_pszMkDocument, out untypedProperty)))
                        {
                            var docPath = (string)untypedProperty;
                            if (!(dataModel.isDoc && (dataModel.currDoc == docPath)))
                            {
                                dataModel.AddWindow(docPath);
                            }
                        }
                        break;
                    case FrameType.ToolWindow:
                        if (ErrorHandler.Succeeded(windowFrame.GetProperty((int)__VSFPROPID.VSFPROPID_Caption, out untypedProperty)))
                        {
                            var caption = (string)untypedProperty;
                            Guid windowId;
                            if (ErrorHandler.Succeeded(windowFrame.GetGuidProperty((int)__VSFPROPID.VSFPROPID_GuidPersistenceSlot, out windowId)))
                            {
                                if (caption != "Macro Explorer")
                                {
                                    if (!((dataModel.isDoc == false) && (dataModel.currWindow == windowId)))
                                    {
                                        dataModel.AddWindow(windowId, caption);
                                    }
                                }
                            }
                        }
                        break;
                }
            }
            else
            {
                return VSConstants.S_OK;
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
            if (monSelCookie != 0U && serviceProvider != null)
            {
                var monSel = (IVsMonitorSelection)serviceProvider.GetService(typeof(SVsShellMonitorSelection));
                if (monSel != null)
                {
                    // NOTE: We can ignore the return code here as there really isn't anything reasonable we could do to deal with failure,
                    // and it is essentially a no-fail method.
                    monSel.UnadviseSelectionEvents(monSelCookie);
                    monSelCookie = 0U;
                }
            }

            this.serviceProvider = null;
        }
    }
}
