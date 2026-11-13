using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace EZOffice
{
    internal sealed class HorizontalScroll : IDisposable
    {
        private const int WhGetMessage = 3;
        private const int PmRemove = 1;
        private const int WmMouseWheel = 0x020A;
        private const int MkShift = 0x0004;
        private const int MkControl = 0x0008;
        private const int VkMenu = 0x12;
        private const int WheelDelta = 120;
        private const int LinesPerNotch = 3;
        private const int RootWindow = 2;

        private readonly string frameClass;
        private readonly string editorClass;
        private readonly Func<IntPtr, int, double, double, bool> scroll;
        private readonly HookProc callback;
        private readonly StringBuilder className = new StringBuilder(32);
        private readonly Control dispatcher = new Control();
        private IntPtr hook;
        private IntPtr lastPane;
        private int wheelRemainder;

        public HorizontalScroll(string frameClass, string editorClass, Func<IntPtr, int, double, double, bool> scroll)
        {
            this.frameClass = frameClass;
            this.editorClass = editorClass;
            this.scroll = scroll;
            callback = OnMessage;
            dispatcher.CreateControl();
            hook = SetWindowsHookEx(WhGetMessage, callback, IntPtr.Zero, GetCurrentThreadId());
            if (hook == IntPtr.Zero)
            {
                int error = Marshal.GetLastWin32Error();
                dispatcher.Dispose();
                throw new Win32Exception(error);
            }
        }

        private IntPtr OnMessage(int code, IntPtr removed, IntPtr data)
        {
            if (hook != IntPtr.Zero && code == 0 && removed == (IntPtr)PmRemove)
            {
                // Read the common MSG prefix. Message is the uint after hwnd.
                if (Marshal.ReadInt32(data, IntPtr.Size) == WmMouseWheel)
                {
                    var message = (WheelMessage)Marshal.PtrToStructure(data, typeof(WheelMessage));
                    if (HandleWheel(message))
                        Marshal.WriteInt32(data, IntPtr.Size, 0); // WM_NULL
                }
            }
            return CallNextHookEx(hook, code, removed, data);
        }

        private bool HandleWheel(WheelMessage message)
        {
            int flagsAndDelta = unchecked((int)message.WParam.ToInt64());
            if ((flagsAndDelta & (MkShift | MkControl)) != MkShift || GetKeyState(VkMenu) < 0)
            {
                wheelRemainder = 0;
                return false;
            }

            // Wheel messages can go to the focused control. Use the pointer to
            // leave scrolling over the ribbon, navigation pane and dialogs alone.
            long position = message.LParam.ToInt64();
            var point = new Point { X = unchecked((short)position), Y = unchecked((short)(position >> 16)) };
            IntPtr pane = FindDocumentPane(WindowFromPoint(point));
            IntPtr frame = GetAncestor(pane, RootWindow);
            if (pane == IntPtr.Zero || frame == IntPtr.Zero || frame != GetActiveWindow()
                || GetAncestor(message.Hwnd, RootWindow) != frame || !IsWindowEnabled(frame)
                || !HasClass(frame, frameClass))
            {
                wheelRemainder = 0;
                return false;
            }

            double editorX = 100;
            double editorY = 0;
            // In recent PowerPoint builds the slide, thumbnails and notes share
            // one native window. Use the view's splitter positions for hit testing.
            if (HasClass(pane, "mdiClass"))
            {
                Rect bounds;
                if (!GetWindowRect(GetAncestor(pane, 1), out bounds) || bounds.Right <= bounds.Left || bounds.Bottom <= bounds.Top)
                    return false;
                editorX = 100.0 * (point.X - bounds.Left) / (bounds.Right - bounds.Left);
                editorY = 100.0 * (point.Y - bounds.Top) / (bounds.Bottom - bounds.Top);
            }

            int delta = unchecked((short)(flagsAndDelta >> 16));
            // Office automation runs after the message hook has returned.
            dispatcher.BeginInvoke((MethodInvoker)delegate
            {
                if (hook == IntPtr.Zero || frame != GetActiveWindow())
                    return;
                int movement = (pane == lastPane ? wheelRemainder : 0) + delta;
                if (scroll(frame, -(movement / WheelDelta) * LinesPerNotch, editorX, editorY))
                {
                    lastPane = pane;
                    wheelRemainder = movement % WheelDelta;
                }
                else
                {
                    wheelRemainder = 0;
                    // A sent message bypasses the queue hook and reaches Office unchanged.
                    SendMessage(message.Hwnd, WmMouseWheel, message.WParam, message.LParam);
                }
            });
            return true;
        }

        private IntPtr FindDocumentPane(IntPtr window)
        {
            while (window != IntPtr.Zero)
            {
                if (HasClass(window, editorClass) || (frameClass == "PPTFrameClass" && HasClass(window, "mdiClass")))
                    return window;
                window = GetAncestor(window, 1); // GA_PARENT excludes owners
            }
            return IntPtr.Zero;
        }

        private bool HasClass(IntPtr window, string expected)
        {
            return GetClassName(window, className, className.Capacity) != 0 && className.ToString() == expected;
        }

        internal static bool IsInModalDialog(string frameClass)
        {
            IntPtr window = GetActiveWindow();
            if (window == IntPtr.Zero)
                return false;
            var name = new StringBuilder(32);
            GetClassName(window, name, name.Capacity);
            return name.ToString() != frameClass || !IsWindowEnabled(window);
        }

        public void Dispose()
        {
            if (hook != IntPtr.Zero)
            {
                UnhookWindowsHookEx(hook);
                hook = IntPtr.Zero;
            }
            dispatcher.Dispose();
        }

        private delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct Point { public int X; public int Y; }

        [StructLayout(LayoutKind.Sequential)]
        private struct Rect { public int Left; public int Top; public int Right; public int Bottom; }

        [StructLayout(LayoutKind.Sequential)]
        private struct WheelMessage
        {
            public IntPtr Hwnd;
            public int Message;
            public IntPtr WParam;
            public IntPtr LParam;
        }

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern IntPtr SetWindowsHookEx(int id, HookProc callback, IntPtr module, uint thread);

        [DllImport("user32.dll")]
        private static extern bool UnhookWindowsHookEx(IntPtr hook);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(Point point);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr window, StringBuilder name, int capacity);

        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr window, int flags);

        [DllImport("user32.dll")]
        private static extern IntPtr GetActiveWindow();

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr window, out Rect bounds);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr window, int message, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowEnabled(IntPtr window);

        [DllImport("user32.dll")]
        private static extern short GetKeyState(int virtualKey);
    }
}
