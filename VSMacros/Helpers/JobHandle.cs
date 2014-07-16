using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Internal.VisualStudio.Shell;

namespace VSMacros.Helpers
{
    internal class JobHandle : SafeHandle
    {
        private readonly ILogger logger;
        private readonly string jobName;

        internal static JobHandle CreateNewJob(ILogger logger)
        {
            Validate.IsNotNull(logger, "logger");

            JobHandle job = new JobHandle(logger);
            if (job.InitializeJob())
            {
                return job;
            }
            else
            {
                return null;
            }
        }

        private JobHandle()
            : base(NativeMethods.INVALIDHANDLEVALUE, ownsHandle: true)
        {
        }

        private JobHandle(ILogger logger)
            : this()
        {
            this.jobName = Guid.NewGuid().ToString();
            this.logger = logger;
        }

        internal bool InitializeJob()
        {

            this.logger.LogInfo("InitializeJob", "JobHandle : Creating job: {0}", this.jobName);

            // Create a job object and add the process to it. We do this to ensure that we can kill the process and 
            // all its children in one shot.  Simply killing the process does not kill its children.
            NativeMethods.SECURITY_ATTRIBUTES jobSecAttributes = new NativeMethods.SECURITY_ATTRIBUTES();
            jobSecAttributes.bInheritHandle = true;
            jobSecAttributes.lpSecurityDescriptor = IntPtr.Zero;
            jobSecAttributes.nLength = Marshal.SizeOf(typeof(NativeMethods.SECURITY_ATTRIBUTES));

            IntPtr pointer = NativeMethods.CreateJobObject(ref jobSecAttributes, this.jobName);
            int createJobObjectWin32Error = Marshal.GetLastWin32Error();

            this.SetHandle(pointer);

            if (this.IsInvalid)
            {
                this.logger.LogError("InitializeJob", "JobHandle : Failed to SetInformation for job '{0}' because '{1}'", this.jobName, createJobObjectWin32Error);
            }
            else
            {
                this.logger.LogInfo("InitializeJob", "JobHandle : Created job object, setting basic information.");

                // LimitFlags include JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
                // to terminate all processes associated with the job when the last job handle is closed
                NativeMethods.JOBOBJECT_BASIC_LIMIT_INFORMATION basicInfo = new NativeMethods.JOBOBJECT_BASIC_LIMIT_INFORMATION();

                NativeMethods.JOBOBJECT_EXTENDED_LIMIT_INFORMATION extendedInfo = new NativeMethods.JOBOBJECT_EXTENDED_LIMIT_INFORMATION { BasicLimitInformation = basicInfo };
                extendedInfo.BasicLimitInformation.LimitFlags = NativeMethods.JOB_OBJECT_LIMIT.KILL_ON_JOB_CLOSE;

                int length = Marshal.SizeOf(typeof(NativeMethods.JOBOBJECT_EXTENDED_LIMIT_INFORMATION));
                IntPtr extendedInfoPtr = Marshal.AllocHGlobal(length);
                Marshal.StructureToPtr(extendedInfo, extendedInfoPtr, false);

                bool result = NativeMethods.SetInformationJobObject(this, NativeMethods.JOBOBJECTINFOCLASS.JobObjectExtendedLimitInformation, ref extendedInfo, length);
                int setInformationJobObjectWin32Error = Marshal.GetLastWin32Error();
                if (!result)
                {
                    this.logger.LogError("InitializeJob", "JobHandle : Failed to SetInformation for job '{0}' because '{1}'", this.jobName, setInformationJobObjectWin32Error);
                }
                else
                {
                    return true;
                }
            }

            return false;
        }

        public bool AddProcess(Process process)
        {
            this.logger.LogInfo("AddProcess", "Adding process '{0}' with PID '{1}' to job '{2}'", process.ProcessName, process.Id, this.jobName);
            using (ProcessHandle processHandle = new ProcessHandle(process.Id))
            {
                if (!processHandle.IsInvalid)
                {
                    bool result = NativeMethods.AssignProcessToJobObject(this, processHandle);
                    int lastWin32Error = Marshal.GetLastWin32Error();
                    if (result)
                    {
                        this.logger.LogInfo("AddProcess", "Process with PID '{0}' added to job '{1}'", process.Id, this.jobName);
                        return true;
                    }

                    this.logger.LogError("AddProcess", "AddProcess : Failed to AssignProcessToJob '{0}' because '{1}'", this.jobName, lastWin32Error);
                }
            }

            return false;
        }

        public bool KillProcessTree()
        {
            return NativeMethods.TerminateJobObject(this, 1);
        }

        protected override bool ReleaseHandle()
        {
            return NativeMethods.CloseHandle(this.handle);
        }

        public override bool IsInvalid
        {
            get { return this.IsClosed || this.handle == NativeMethods.INVALIDHANDLEVALUE; }
        }

        private class ProcessHandle : SafeHandle
        {
            private ProcessHandle()
                : base(NativeMethods.INVALIDHANDLEVALUE, ownsHandle: true)
            {
            }

            public ProcessHandle(int processId)
                : this()
            {
                this.SetHandle(NativeMethods.OpenProcess(NativeMethods.PROCESSALLACCESS, false, (uint)processId));
            }

            protected override bool ReleaseHandle()
            {
                return NativeMethods.CloseHandle(this.handle);
            }

            public override bool IsInvalid
            {
                get { return this.IsClosed || this.handle == NativeMethods.INVALIDHANDLEVALUE; }
            }
        }
    }
}