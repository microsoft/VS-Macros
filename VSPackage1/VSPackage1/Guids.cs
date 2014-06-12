// Guids.cs
// MUST match guids.h
using System;

namespace Company.VSPackage1
{
    static class GuidList
    {
        public const string guidVSPackage1PkgString = "a6687369-228e-483b-bdd7-b5b714a5ba74";
        public const string guidVSPackage1CmdSetString = "046cdb5d-d43f-45d5-83a4-b39e3c221e42";

        public static readonly Guid guidVSPackage1CmdSet = new Guid(guidVSPackage1CmdSetString);
    };
}