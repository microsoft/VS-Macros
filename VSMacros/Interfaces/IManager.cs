//-----------------------------------------------------------------------
// <copyright file="IManager.cs" company="Microsoft Corporation">
//     Copyright Microsoft Corporation. All rights reserved.
// </copyright>
//-----------------------------------------------------------------------

namespace VSMacros.Interfaces
{
    public interface IManager
    {
        void StartRecording();

        void StopRecording();

        void Playback(string path, int times);

        void StopPlayback();

        void SaveCurrent();

        void Refresh();

        void Edit();

        void Rename();

        void AssignShortcut();

        void Delete();

        void Close();
    }
}
