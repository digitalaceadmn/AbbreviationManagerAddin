using System;
using System.Windows.Forms;

namespace AbbreviationWordAddin
{
    public partial class ReplaceDialog : Form
    {
        public enum ReplaceAction
        {
            Replace,
            ReplaceAll,
            Ignore,
            IgnoreAll,
            Cancel,
            Close
        }

        public ReplaceAction UserChoice { get; private set; }

        // Custom constructor – pass phrase & replacement
        public ReplaceDialog(string phrase, string replacement)
        {
            InitializeComponent();

            // set textbox values
            txtPhrase.Text = phrase;
            txtReplacement.Text = replacement;

            // wire up button click events
            btnReplace.Click += (s, e) =>
            {
                UserChoice = ReplaceAction.Replace;
                DialogResult = DialogResult.OK;
            };

            btnReplaceAll.Click += (s, e) =>
            {
                UserChoice = ReplaceAction.ReplaceAll;
                DialogResult = DialogResult.OK;
            };

            btnIgnore.Click += (s, e) =>
            {
                UserChoice = ReplaceAction.Ignore;
                DialogResult = DialogResult.OK;
            };

            btnIgnoreAll.Click += (s, e) =>
            {
                UserChoice = ReplaceAction.IgnoreAll;
                DialogResult = DialogResult.OK;
            };

            btnCancel.Click += (s, e) =>
            {
                UserChoice = ReplaceAction.Cancel;
                DialogResult = DialogResult.Cancel;
            };

            btnClose.Click += (s, e) =>
            {
                UserChoice = ReplaceAction.Close;
                DialogResult = DialogResult.Cancel;
            };
        }

        private void ReplaceDialog_Load(object sender, EventArgs e) { }
        private void label1_Click(object sender, EventArgs e) { }
    }
}
