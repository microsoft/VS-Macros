//-----------------------------------------------------------------------
// <copyright file="IRecorder.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace VSMacros.Interfaces
{
    interface IRecorder
    {
        void StartRecording();
        void StopRecording(string path);
    }
}
