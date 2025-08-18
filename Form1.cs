using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AbbreviationWordAddin
{
    public partial class Form1 : Form
    {
        [DllImport("user32.dll")]
        private static extern bool HideCaret(IntPtr hWnd);

        public Form1()
        {
            InitializeComponent();

            // Hide caret whenever the RichTextBox is focused or changed
            richTextBox1.GotFocus += (s, e) => HideCaret(richTextBox1.Handle);
            richTextBox1.SelectionChanged += (s, e) => HideCaret(richTextBox1.Handle);
            richTextBox1.MouseDown += (s, e) => HideCaret(richTextBox1.Handle);
            richTextBox1.KeyUp += (s, e) => HideCaret(richTextBox1.Handle);
        }

        public void AppendLog(string text)
        {
            richTextBox1.AppendText(text + Environment.NewLine);
        }
    }
}
