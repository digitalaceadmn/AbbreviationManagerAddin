using System;
using System.Drawing;
using System.Windows.Forms;

namespace AbbreviationWordAddin
{
    public partial class SuggestionPaneControl : UserControl
    {
        private System.ComponentModel.IContainer components = null;

        private TextBox textBoxInput;
        private TabControl tabControlModes;
        private TabPage tabPageAbbreviation;
        private TabPage tabPageReverse;
        private TabPage tabPageDictionary;
        private ListView listViewAbbrev;
        private ListView listViewReverse;
        private ListView listViewDictionary;

        private Label lblWord;
        private Label lblReplacement;
        private TextBox txtWord;
        private TextBox txtReplacement;

        private Button btnReplace;
        private Button btnReplaceAll;
        private Button btnIgnore;
        private Button btnIgnoreAll;

        private FlowLayoutPanel buttonPanel;
        private TableLayoutPanel bottomLayout;
        private TableLayoutPanel mainLayout;

        protected override void Dispose(bool disposing)
        {
            if (disposing && components != null)
                components.Dispose();
            base.Dispose(disposing);
        }

 

        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();

            textBoxInput = new TextBox();
            tabControlModes = new TabControl();
            tabPageAbbreviation = new TabPage("Abbreviations");
            tabPageReverse = new TabPage("Reverse Abbreviations");
            tabPageDictionary = new TabPage("Dictionary");

            listViewAbbrev = new ListView();
            listViewReverse = new ListView();
            listViewDictionary = new ListView();

            lblWord = new Label();
            lblReplacement = new Label();
            txtWord = new TextBox();
            txtReplacement = new TextBox();

            btnReplace = new Button();
            btnReplaceAll = new Button();
            btnIgnore = new Button();
            btnIgnoreAll = new Button();

            buttonPanel = new FlowLayoutPanel();
            bottomLayout = new TableLayoutPanel();
            mainLayout = new TableLayoutPanel();

            SuspendLayout();

            // ================= SEARCH INPUT =================
            textBoxInput.Dock = DockStyle.Fill;
            textBoxInput.Margin = new Padding(10);
            textBoxInput.Height = 32;

            // ================= LIST VIEWS =================
            ConfigureListView(listViewAbbrev);
            ConfigureListView(listViewReverse);
            ConfigureListView(listViewDictionary);

            tabPageAbbreviation.Controls.Add(listViewAbbrev);
            tabPageReverse.Controls.Add(listViewReverse);
            tabPageDictionary.Controls.Add(listViewDictionary);

            tabControlModes.Dock = DockStyle.Fill;
            tabControlModes.Controls.Add(tabPageAbbreviation);
            tabControlModes.Controls.Add(tabPageReverse);
            tabControlModes.Controls.Add(tabPageDictionary);

            lblWord.Text = "Word / Phrase";
            lblReplacement.Text = "Replacement";
            lblWord.AutoSize = true;
            lblReplacement.AutoSize = true;

            txtWord.Dock = DockStyle.Fill;
            txtReplacement.Dock = DockStyle.Fill;

            buttonPanel.FlowDirection = FlowDirection.LeftToRight;
            buttonPanel.WrapContents = false;
            buttonPanel.AutoSize = true;
            buttonPanel.Anchor = AnchorStyles.Right;

            bottomLayout.ColumnCount = 3;
            bottomLayout.RowCount = 2;
            bottomLayout.Dock = DockStyle.Fill;
            bottomLayout.Padding = new Padding(10);
            bottomLayout.AutoSize = true;

            bottomLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            bottomLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            bottomLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            bottomLayout.Controls.Add(lblWord, 0, 0);
            bottomLayout.Controls.Add(txtWord, 1, 0);
            bottomLayout.Controls.Add(lblReplacement, 0, 1);
            bottomLayout.Controls.Add(txtReplacement, 1, 1);
            bottomLayout.Controls.Add(buttonPanel, 2, 0);
            bottomLayout.SetRowSpan(buttonPanel, 2);

            // ================= MAIN LAYOUT =================
            mainLayout.ColumnCount = 1;
            mainLayout.RowCount = 3;
            mainLayout.Dock = DockStyle.Fill;

            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            mainLayout.Controls.Add(textBoxInput, 0, 0);
            mainLayout.Controls.Add(bottomLayout, 0, 1);
            mainLayout.Controls.Add(tabControlModes, 0, 2);

            Controls.Add(mainLayout);

            Name = "SuggestionPaneControl";
            Size = new Size(1200, 900);

            ResumeLayout(false);
        }

        // ================= ON LOAD =================
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            ApplyButtonStyle(btnReplace, "Replace");
            ApplyButtonStyle(btnReplaceAll, "Replace All");
            ApplyButtonStyle(btnIgnore, "Ignore");
            ApplyButtonStyle(btnIgnoreAll, "Ignore All");

            buttonPanel.Controls.Add(btnReplace);
            buttonPanel.Controls.Add(btnReplaceAll);
            buttonPanel.Controls.Add(btnIgnore);
            buttonPanel.Controls.Add(btnIgnoreAll);

            // Event wiring
            textBoxInput.TextChanged += textBoxInput_TextChanged;
            btnReplace.Click += btnReplace_Click;
            btnReplaceAll.Click += btnReplaceAll_Click;
            btnIgnore.Click += btnIgnore_Click;
            btnIgnoreAll.Click += btnIgnoreAll_Click;

            KeyDown += SuggestionPaneControl_KeyDown;
        }

        // ================= STYLES =================
        private void ApplyButtonStyle(Button btn, string text)
        {
            btn.Text = text;
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
            btn.BackColor = Color.FromArgb(240, 242, 245);
            btn.ForeColor = Color.Black;
            btn.Width = 130;
            btn.Height = 32;
            btn.Margin = new Padding(6, 0, 0, 0);
            btn.TextAlign = ContentAlignment.MiddleCenter;
        }

        private void ConfigureListView(ListView lv)
        {
            lv.Dock = DockStyle.Fill;
            lv.FullRowSelect = true;
            lv.HideSelection = false;
            lv.View = View.Details;
        }

        // ================= KEYBOARD SHORTCUTS =================
        private void SuggestionPaneControl_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.R)
            {
                btnReplace.PerformClick();
                e.Handled = true;
            }
            else if (e.Control && e.Shift && e.KeyCode == Keys.R)
            {
                btnReplaceAll.PerformClick();
                e.Handled = true;
            }
        }

    }
}
