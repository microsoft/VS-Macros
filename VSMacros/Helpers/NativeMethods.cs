//-----------------------------------------------------------------------
// <copyright file="NativeMethods.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace VSMacros.Helpers
{
    internal static class NativeMethods
    {
        public const int S_OK = 0x00000000;
        public const int ROTFLAGS_REGISTRATIONKEEPSALIVE = 1;
        public const int ROTFLAGS_ALLOWANYCLIENT = 2;

        [DllImport("ole32.dll")]
        internal static extern int CreateItemMoniker([MarshalAs(UnmanagedType.LPWStr)] string
           lpszDelim, [MarshalAs(UnmanagedType.LPWStr)] string lpszItem,
           out System.Runtime.InteropServices.ComTypes.IMoniker ppmk);

        [DllImport("ole32.dll")]
        public static extern int GetRunningObjectTable(int reserved,
                                  out IRunningObjectTable prot);

        [DllImport("ole32.dll")]
        public static extern int CreateBindCtx(int reserved,
                                  out IBindCtx bc);

                public const uint PROCESSALLACCESS = 0x1F0FFF; // All possible access rights for a process object.
        public const int ERROR_ALREADY_EXISTS = 183;
        public static readonly IntPtr INVALIDHANDLEVALUE = new IntPtr(-1);
        
        [StructLayout(LayoutKind.Sequential)]
        public struct SECURITY_ATTRIBUTES
        {
            public int nLength;
            public IntPtr lpSecurityDescriptor;
            public bool bInheritHandle;
        }

        public enum JOB_OBJECT_LIMIT : uint
        {
            // Basic limits
            WORKINGSET = 0x00000001,
            PROCESS_TIME = 0x00000002,
            JOB_TIME = 0x00000004,
            ACTIVE_PROCESS = 0x00000008,
            AFFINITY = 0x00000010,
            PRIORITY_CLASS = 0x00000020,
            PRESERVE_JOB_TIME = 0x00000040,
            SCHEDULING_CLASS = 0x00000080,
            
            // Extended limits
            PROCESS_MEMORY = 0x00000100,
            JOB_MEMORY = 0x00000200,
            DIE_ON_UNHANDLED_EXCEPTION = 0x00000400,
            BREAKAWAY_OK = 0x00000800,
            SILENT_BREAKAWAY_OK = 0x00001000,
            KILL_ON_JOB_CLOSE = 0x00002000,
            SUBSET_AFFINITY = 0x00004000,
        }

        public struct IO_COUNTERS
        {
            public ulong ReadOperationCount;
            public ulong WriteOperationCount;
            public ulong OtherOperationCount;
            public ulong ReadTransferCount;
            public ulong WriteTransferCount;
            public ulong OtherTransferCount;
        }

        public struct JOBOBJECT_EXTENDED_LIMIT_INFORMATION
        {
            public JOBOBJECT_BASIC_LIMIT_INFORMATION BasicLimitInformation;
            public IO_COUNTERS IoInfo;
            public UIntPtr ProcessMemoryLimit;
            public UIntPtr JobMemoryLimit;
            public UIntPtr PeakProcessMemoryUsed;
            public UIntPtr PeakJobMemoryUsed;
        }

        public struct JOBOBJECT_BASIC_LIMIT_INFORMATION
        {
            public long PerProcessUserTimeLimit;
            public long PerJobUserTimeLimit;
            public JOB_OBJECT_LIMIT LimitFlags;
            public UIntPtr MinimumWorkingSetSize;
            public UIntPtr MaximumWorkingSetSize;
            public uint ActiveProcessLimit;
            public UIntPtr Affinity;
            public uint PriorityClass;
            public uint SchedulingClass;
        }

        public enum JOBOBJECTINFOCLASS
        {
            JobObjectBasicAccountingInformation = 1,
            JobObjectBasicLimitInformation,
            JobObjectBasicProcessIdList,
            JobObjectBasicUIRestrictions,
            JobObjectSecurityLimitInformation,
            JobObjectEndOfJobTimeInformation,
            JobObjectAssociateCompletionPortInformation,
            JobObjectBasicAndIoAccountingInformation,
            JobObjectExtendedLimitInformation,
            JobObjectJobSetInformation,
            MaxJobObjectInfoClass
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr CreateJobObject([In] ref SECURITY_ATTRIBUTES lpJobAttributes,
            string lpName);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetInformationJobObject(
            SafeHandle hJob,
            JOBOBJECTINFOCLASS JobObjectInfoClass,
            ref JOBOBJECT_EXTENDED_LIMIT_INFORMATION lpJobObjectInfo,
            int cbJobObjectInfoLength);

        [DllImport("kernel32.dll")]
        public static extern bool AssignProcessToJobObject(SafeHandle hJob, SafeHandle hProcess);

        [DllImport("kernel32.dll")]
        public static extern bool TerminateJobObject(SafeHandle hJob, uint uExitCode);

        [DllImport("kernel32", SetLastError = true)]
        public static extern bool CloseHandle(IntPtr handle);

        [DllImport("kernel32.dll")]
        public static extern IntPtr OpenProcess(uint dwDesiredAccess, bool bInheritHandle,
           uint dwProcessId);
    }
}
