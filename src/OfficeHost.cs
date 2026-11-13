using System;
using System.Diagnostics;

namespace EZOffice
{
    internal sealed class OfficeHost
    {
        private readonly dynamic application;
        public string Name { get; private set; }
        public string FrameClass { get; private set; }
        public string EditorClass { get; private set; }

        public OfficeHost(object application)
        {
            this.application = application;
            using (Process process = Process.GetCurrentProcess())
            {
                switch (process.ProcessName.ToUpperInvariant())
                {
                    case "WINWORD": Name = "Word"; FrameClass = "OpusApp"; EditorClass = "_WwG"; break;
                    case "EXCEL": Name = "Excel"; FrameClass = "XLMAIN"; EditorClass = "EXCEL7"; break;
                    case "POWERPNT": Name = "PowerPoint"; FrameClass = "PPTFrameClass"; EditorClass = "paneClassDC"; break;
                    default: throw new NotSupportedException("EZOffice supports Word, Excel and PowerPoint.");
                }
            }
        }

        public object GetFilesReadyForSaving()
        {
            switch (Name)
            {
                case "Word": return (int)application.BackgroundSavingStatus == 0 ? application.Documents : null;
                case "Excel": return (bool)application.Ready ? application.Workbooks : null;
                default: return application.Presentations;
            }
        }

        public bool TryScroll(IntPtr frame, int lines, double editorX, double editorY)
        {
            dynamic window = application.ActiveWindow;
            if (window == null)
                return false;

            if (Name == "PowerPoint")
            {
                int view = (int)window.ViewType;
                const int normalView = 9;
                if (view == normalView && (editorX < (float)window.SplitHorizontal || editorY > (float)window.SplitVertical))
                    return false;
                if (view != normalView && view != 1 && view != 2 && view != 8)
                    return false;
            }
            else if ((IntPtr)(int)window.Hwnd != frame)
            {
                return false;
            }

            if (lines != 0)
                window.SmallScroll(0, 0, Math.Max(0, lines), Math.Max(0, -lines));
            return true;
        }
    }
}
