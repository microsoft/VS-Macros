using System;
using System.IO;

namespace VSMacros.Engines
{
    interface IManager
    {
        void ToggleRecording();

        void Playback(string path, int times);

        void StopPlayback();

        void SaveCurrent();
    }

    private sealed class Manager : IManager
    {
        static readonly Manager _instance = new Manager();

        private Manager() 
        { 
        }

        public static Manager Instance
        {
            get { return _instance; }
        }

        public void ToggleRecording()
        {
        }

        public void Playback(string path, int times) 
        {
        }

        public void StopPlayback() 
        {
        }

        public void SaveCurrent() 
        {
        }

        private FileStream LoadFile(string path) 
        { 
            return null; 
        }

        private void SaveToFile(Stream str)
        { 
        }
    }
}
