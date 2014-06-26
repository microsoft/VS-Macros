using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VSMacros.Interfaces
{
    interface IManager
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
