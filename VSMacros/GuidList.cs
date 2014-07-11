//-----------------------------------------------------------------------
// <copyright file="GuidList.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;

namespace VSMacros
{
    internal static class GuidList
    {
        public const string GuidVSMacrosPkgString = "17f25bfa-a812-4b6a-a07e-7f73aa975f8b";
        public const string GuidVSMacrosCmdSetString = "ad5b992c-e71a-4619-8d95-502b3b8de2c1";
        public const string GuidToolWindowPersistanceString = "56fbfa32-c049-4fd5-9b54-39fcdf33629d";

        public static readonly Guid GuidVSMacrosCmdSet = new Guid(GuidVSMacrosCmdSetString);
    }
}