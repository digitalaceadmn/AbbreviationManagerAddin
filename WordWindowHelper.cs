using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public static class WordWindowHelper
{
    [DllImport("user32.dll")]
    private static extern IntPtr GetActiveWindow();

    public static IWin32Window GetWordWindow()
    {
        return new WindowWrapper(GetActiveWindow());
    }

    private class WindowWrapper : IWin32Window
    {
        private readonly IntPtr _handle;
        public WindowWrapper(IntPtr handle)
        {
            _handle = handle;
        }

        public IntPtr Handle => _handle;
    }
}
