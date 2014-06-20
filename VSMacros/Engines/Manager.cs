using System;
using System.IO;

namespace VSMacros.Engines
{
    public sealed class Manager
    {
        static readonly Manager _instance = new Manager();

        private Manager() { }

        public static Manager Instance
        {
            get { return _instance; }
        }

        #region Command Handling

        public void StartRecording() { }

        public void StopRecording() { }

        public void Playback(string path, int times) { }

        public void StopPlayback() { }

        public void SaveCurrent() { }

        private FileStream LoadFile(string path) { return null; }
        private void SaveToFile(Stream str) { }
        #endregion
    }
}
