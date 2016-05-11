//-----------------------------------------------------------------------
// <copyright file="FilePath.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System;

namespace VisualStudio.Macros.ExecutionEngine.Pipes
{
    [Serializable]
    public class FilePath
    {
        public int Iterations;
        public string Path;
    }
}
