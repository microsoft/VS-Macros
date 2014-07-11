//-----------------------------------------------------------------------
// <copyright file="IRecorder.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

using System.IO;
namespace VSMacros.Interfaces
{
    interface IRecorder
    {
        void StartRecording();

        void StopRecording(string path);
    }
}
