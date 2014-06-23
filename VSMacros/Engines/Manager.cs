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

        public void ToggleRecording() { }

        // TODO adjust function call when the executer interface is written
        public void Playback(string path, int times) { }

        public void StopPlayback() { }

        public void SaveCurrent() { }
        public void LoadCurrent() { }

        private FileStream LoadFile(string path) { return null; }
        private void SaveToFile(Stream str) { }

    }
}
