//-----------------------------------------------------------------------
// <copyright file="CommandHelper.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Internal.VisualStudio.Shell;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell.Interop;

namespace ExecutionEngine.Helpers
{
    public sealed class CommandHelper : ICommandHelper
    {
        private readonly IOleCommandTarget shellCmdTarget;
        private readonly IVsCmdNameMapping cmdNameMapping;

        private System.Runtime.InteropServices.ComTypes.IRunningObjectTable GetRunningObjectTable()
        {
            System.Runtime.InteropServices.ComTypes.IRunningObjectTable rot;
            int hr = NativeMethods.GetRunningObjectTable(0, out rot);
            if (ErrorHandler.Failed(hr))
            {
                ErrorHandler.ThrowOnFailure(hr, null);
            }

            return rot;
        }

        private System.Runtime.InteropServices.ComTypes.IMoniker GetCmdDispatcherMoniker()
        {
            System.Runtime.InteropServices.ComTypes.IMoniker moniker;
            string delim = "!";
            Guid guid = Marshal.GenerateGuidForType(typeof(SUIHostCommandDispatcher));
            string progID = string.Format(CultureInfo.InvariantCulture, "{0:B}", guid);
            int hResult = NativeMethods.CreateItemMoniker(delim, progID, out moniker);

            return moniker;
        }

        private System.Runtime.InteropServices.ComTypes.IMoniker GetCmdNameMoniker()
        {
            System.Runtime.InteropServices.ComTypes.IMoniker moniker;
            string delim = "!";
            Guid guid = Marshal.GenerateGuidForType(typeof(SVsCmdNameMapping));
            string progID = string.Format(CultureInfo.InvariantCulture, "{0:B}", guid);
            int hResult = NativeMethods.CreateItemMoniker(delim, progID, out moniker);

            return moniker;
        }

        private object GetObjectFromRot(System.Runtime.InteropServices.ComTypes.IRunningObjectTable rot, System.Runtime.InteropServices.ComTypes.IMoniker moniker)
        {
            object objectFromRot;
            int hr = rot.GetObject(moniker, out objectFromRot);
            if (ErrorHandler.Failed(hr))
            {
                ErrorHandler.ThrowOnFailure(hr, null);
            }

            return objectFromRot;
        }

        public CommandHelper(System.IServiceProvider serviceProvider)
        {
            var rot = this.GetRunningObjectTable();

            System.Runtime.InteropServices.ComTypes.IMoniker cmdDispatchMoniker = this.GetCmdDispatcherMoniker();
            System.Runtime.InteropServices.ComTypes.IMoniker cmdNameMoniker = this.GetCmdNameMoniker();

            var cmdDispatchObject = this.GetObjectFromRot(rot, cmdDispatchMoniker);
            var cmdNameObject = this.GetObjectFromRot(rot, cmdNameMoniker);

            if (serviceProvider == null)
            {
                throw new ArgumentNullException("serviceProvider");
            }

            this.shellCmdTarget = (IOleCommandTarget)cmdDispatchObject;
            this.cmdNameMapping = (IVsCmdNameMapping)cmdNameObject;
        }   

        public void DispatchCommandByName(object canonicalName)
        {
            Guid cmdSet;
            uint cmdId;
            ErrorHandler.ThrowOnFailure(this.cmdNameMapping.MapNameToGUIDID((string)canonicalName, out cmdSet, out cmdId));
            this.DispatchCommandHelper(cmdSet, cmdId);
        }

        public void DispatchCommand(object commandSet, object commandId)
        {
            this.DispatchCommandHelper(commandSet, commandId);
        }

        public void DispatchCommandWithArgs(object commandSet, object commandId, ref object pvaIn)
        {
            Validate.IsNotNull(commandSet, "commandSet");
            Validate.IsNotNull(commandId, "commandId");
            Validate.IsNotNull(pvaIn, "pvaIn");

            char character = ((string)pvaIn)[0];
            this.DispatchCommandHelper(commandSet, commandId, () => 
            {
                IntPtr commandPtr = Marshal.AllocHGlobal(16);
                Marshal.GetNativeVariantForObject((ushort)character, commandPtr);
                return commandPtr;
            },
            (ptr) => 
            { 
                if (ptr != IntPtr.Zero) 
                { 
                    Marshal.FreeHGlobal(ptr); 
                } 
            });
        }

        public void ShowMessage(object message)
        {
            Validate.IsNotNull(message, "message");
            MessageBox.Show(message.ToString());
        }

        private void DispatchCommandHelper(object commandSet, object commandId)
        {
            this.DispatchCommandHelper(commandSet, commandId, getCmdArgs: null, releaseCmdArgs: null);
        }

        private void DispatchCommandHelper(Guid commandSet, uint commandId)
        {
            this.DispatchCommandHelper(commandSet, commandId, getCmdArgs: null, releaseCmdArgs: null);
        }

        private void DispatchCommandHelper(object commandSet, object commandId, Func<IntPtr> getCmdArgs, Action<IntPtr> releaseCmdArgs)
        {
            Guid cmdSet = Guid.Parse((string)commandSet);

            uint cmdId = (uint)(int)commandId;

            this.DispatchCommandHelper(cmdSet, cmdId, getCmdArgs, releaseCmdArgs);
        }

        private void DispatchCommandHelper(Guid commandSet, uint commandId, Func<IntPtr> getCmdArgs, Action<IntPtr> releaseCmdArgs)
        {
            if (this.IsCommandEnabled(ref commandSet, commandId))
            {
                IntPtr cmdArgs = IntPtr.Zero;
                try
                {
                    cmdArgs = (getCmdArgs != null) ? getCmdArgs() : IntPtr.Zero;
                    this.shellCmdTarget.Exec(ref commandSet, commandId, (uint)Microsoft.VisualStudio.OLE.Interop.Constants.MSOCMDEXECOPT_DODEFAULT, cmdArgs, IntPtr.Zero);
                }
                finally
                {
                    if (releaseCmdArgs != null && cmdArgs != IntPtr.Zero)
                    {
                        releaseCmdArgs(cmdArgs);
                    }
                }
            }
        }

        private bool IsCommandEnabled(ref Guid cmdSet, uint cmdId)
        {
            OLECMD[] cmds = new OLECMD[] { new OLECMD() { cmdID = cmdId } };
            ErrorHandler.ThrowOnFailure(this.shellCmdTarget.QueryStatus(ref cmdSet, (uint)cmds.Length, cmds, IntPtr.Zero));
            OLECMDF flags = (OLECMDF)cmds[0].cmdf;
            return ((flags & OLECMDF.OLECMDF_ENABLED) == OLECMDF.OLECMDF_ENABLED);
        }

        private ushort GetUShortForIncomingArg(object pvaIn)
        {
            Type type = pvaIn.GetType();
            if (type == typeof(double))
            {
                return (ushort)(double)pvaIn;
            }
            else if (type == typeof(string))
            {
                return (ushort)char.Parse((string)pvaIn);
            }

            throw new ArgumentException(string.Format("Uxpected incoming argument type: {0}", type.FullName));
        }
    }
}
