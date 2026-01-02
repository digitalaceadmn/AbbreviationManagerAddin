using AbbreviationWordAddin;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

public static class WordDialogHelper
{
    public static void ShowInfo(string message, string title = "Abbreviation Tool")
    {
        MessageBox.Show(
            GetWordWindow(),
            message,
            title,
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        );
    }

    public static IWin32Window GetWordWindow()
    {
        try
        {
            Word.Application app = Globals.ThisAddIn.Application;
            Word.Window window = app.ActiveWindow;

            if (window != null)
            {
                return new WindowWrapper(new IntPtr(window.Hwnd));
            }
        }
        catch
        {

        }

        return null;
    }

    private class WindowWrapper : IWin32Window
    {
        public WindowWrapper(IntPtr handle)
        {
            Handle = handle;
        }

        public IntPtr Handle { get; }

        ~WindowWrapper()
        {
            // prevent COM leaks
            if (Handle != IntPtr.Zero)
                Marshal.Release(Handle);
        }
    }
}
