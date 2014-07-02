//-----------------------------------------------------------------------
// <copyright file="PkgCmdIDList.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;

namespace VSMacros
{
    internal static class PkgCmdIDList
    {

        public const uint CmdIdMacroExplorer = 0x101;

        // Tool Window Toolbar
        public const int MacrosToolWindowToolbar = 0x1010;

        // Commands
        public const int CmdIdRecord = 0x2000;
        public const int CmdIdPlayback = 0x2001;
        public const int CmdIdPlaybackMultipleTimes = 0x2002;
        public const int CmdIdSaveTemporaryMacro = 0x2003;
        public const int CmdIdRefresh = 0x2010;
        public const int CmdIdOpenDirectory = 0x2011;
    };
}