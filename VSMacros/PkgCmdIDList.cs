using System;

namespace VSMacros
{
    internal static class PkgCmdIDList
    {
        // Tool Window Toolbar
        public const int MacrosToolWindowToolbar = 0x1050;

        // Context Menus
        public const int BrowserContextMenu = 0x1005;
        public const int CurrentContextMenu = 0x1006;
        public const int MacroContextMenu = 0x1007;
        public const int FolderContextMenu = 0x1008;

        // Commands
        public const int CmdIdMacroExplorer = 0x101;

        public const int CmdIdRecord = 0x2000;
        public const int CmdIdPlayback = 0x2001;
        public const int CmdIdPlaybackMultipleTimes = 0x2002;
        public const int CmdIdSaveTemporaryMacro = 0x2003;
        public const int CmdIdRefresh = 0x2010;
        public const int CmdIdOpenDirectory = 0x2011;

        public const int CmdIdEdit = 0x2020;
        public const int CmdIdRename = 0x2021;
        public const int CmdIdAssignShortcut = 0x2022;
        public const int CmdIdDelete = 0x2023;
    }
}