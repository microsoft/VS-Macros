// IManager.cs


namespace VSMacros.Interfaces
{
    public interface IManager
    {
        void ToggleRecording();

        void Playback(string path, int times);

        void StopPlayback();

        void SaveCurrent();

        void Refresh();

        void Edit();

        void Rename();

        void AssignShortcut();

        void Delete();
    }
}
