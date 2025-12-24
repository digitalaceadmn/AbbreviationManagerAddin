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
        private ListView listViewAbbrev;
        private TabPage tabPageReverse;
        private ListView listViewReverse;
        private TabPage tabPageDictionary;
        private ListView listViewDictionary;

        private Label lblWord;
        private Label lblReplacement;
        private TextBox txtWord;
        private TextBox txtReplacement;

        private Button btnReplace;
        private Button btnReplaceAll;
        private Button btnIgnore;
        private Button btnIgnoreAll;
        private Button btnCancel;
        private Button btnClose;

        private ContextMenuStrip contextMenuStrip1;

        // NEW LAYOUT CONTAINERS
        private TableLayoutPanel mainLayout;
        private Panel bottomPanel;
        private FlowLayoutPanel buttonPanel;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }



        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.textBoxInput = new System.Windows.Forms.TextBox();
            this.tabControlModes = new System.Windows.Forms.TabControl();
            this.tabPageAbbreviation = new System.Windows.Forms.TabPage();
            this.listViewAbbrev = new System.Windows.Forms.ListView();
            this.tabPageReverse = new System.Windows.Forms.TabPage();
            this.listViewReverse = new System.Windows.Forms.ListView();
            this.tabPageDictionary = new System.Windows.Forms.TabPage();
            this.listViewDictionary = new System.Windows.Forms.ListView();
            this.lblWord = new System.Windows.Forms.Label();
            this.lblReplacement = new System.Windows.Forms.Label();
            this.txtWord = new System.Windows.Forms.TextBox();
            this.txtReplacement = new System.Windows.Forms.TextBox();
            this.btnReplace = new System.Windows.Forms.Button();
            this.btnReplaceAll = new System.Windows.Forms.Button();
            this.btnIgnore = new System.Windows.Forms.Button();
            this.btnIgnoreAll = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.mainLayout = new System.Windows.Forms.TableLayoutPanel();
            this.bottomPanel = new System.Windows.Forms.Panel();
            this.buttonPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.tabControlModes.SuspendLayout();
            this.tabPageAbbreviation.SuspendLayout();
            this.tabPageReverse.SuspendLayout();
            this.tabPageDictionary.SuspendLayout();
            this.mainLayout.SuspendLayout();
            this.bottomPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // textBoxInput
            // 
            this.textBoxInput.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBoxInput.Location = new System.Drawing.Point(10, 10);
            this.textBoxInput.Margin = new System.Windows.Forms.Padding(10);
            this.textBoxInput.Name = "textBoxInput";
            this.textBoxInput.Size = new System.Drawing.Size(130, 20);
            this.textBoxInput.TabIndex = 0;
            // 
            // tabControlModes
            // 
            this.tabControlModes.Controls.Add(this.tabPageAbbreviation);
            this.tabControlModes.Controls.Add(this.tabPageReverse);
            this.tabControlModes.Controls.Add(this.tabPageDictionary);
            this.tabControlModes.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControlModes.Location = new System.Drawing.Point(3, 43);
            this.tabControlModes.Name = "tabControlModes";
            this.tabControlModes.SelectedIndex = 0;
            this.tabControlModes.Size = new System.Drawing.Size(144, 1);
            this.tabControlModes.TabIndex = 1;
            // 
            // tabPageAbbreviation
            // 
            this.tabPageAbbreviation.Controls.Add(this.listViewAbbrev);
            this.tabPageAbbreviation.Location = new System.Drawing.Point(4, 22);
            this.tabPageAbbreviation.Name = "tabPageAbbreviation";
            this.tabPageAbbreviation.Size = new System.Drawing.Size(136, 0);
            this.tabPageAbbreviation.TabIndex = 0;
            this.tabPageAbbreviation.Text = "Abbreviations";
            // 
            // listViewAbbrev
            // 
            this.listViewAbbrev.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listViewAbbrev.FullRowSelect = true;
            this.listViewAbbrev.HideSelection = false;
            this.listViewAbbrev.Location = new System.Drawing.Point(0, 0);
            this.listViewAbbrev.Name = "listViewAbbrev";
            this.listViewAbbrev.Size = new System.Drawing.Size(136, 0);
            this.listViewAbbrev.TabIndex = 0;
            this.listViewAbbrev.UseCompatibleStateImageBehavior = false;
            // 
            // tabPageReverse
            // 
            this.tabPageReverse.Controls.Add(this.listViewReverse);
            this.tabPageReverse.Location = new System.Drawing.Point(4, 22);
            this.tabPageReverse.Name = "tabPageReverse";
            this.tabPageReverse.Size = new System.Drawing.Size(192, 0);
            this.tabPageReverse.TabIndex = 1;
            this.tabPageReverse.Text = "Reverse Abbreviations";
            // 
            // listViewReverse
            // 
            this.listViewReverse.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listViewReverse.FullRowSelect = true;
            this.listViewReverse.HideSelection = false;
            this.listViewReverse.Location = new System.Drawing.Point(0, 0);
            this.listViewReverse.Name = "listViewReverse";
            this.listViewReverse.Size = new System.Drawing.Size(192, 0);
            this.listViewReverse.TabIndex = 0;
            this.listViewReverse.UseCompatibleStateImageBehavior = false;
            // 
            // tabPageDictionary
            // 
            this.tabPageDictionary.Controls.Add(this.listViewDictionary);
            this.tabPageDictionary.Location = new System.Drawing.Point(4, 22);
            this.tabPageDictionary.Name = "tabPageDictionary";
            this.tabPageDictionary.Size = new System.Drawing.Size(192, 0);
            this.tabPageDictionary.TabIndex = 2;
            this.tabPageDictionary.Text = "Dictionary";
            // 
            // listViewDictionary
            // 
            this.listViewDictionary.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listViewDictionary.FullRowSelect = true;
            this.listViewDictionary.HideSelection = false;
            this.listViewDictionary.Location = new System.Drawing.Point(0, 0);
            this.listViewDictionary.Name = "listViewDictionary";
            this.listViewDictionary.Size = new System.Drawing.Size(192, 0);
            this.listViewDictionary.TabIndex = 0;
            this.listViewDictionary.UseCompatibleStateImageBehavior = false;
            // 
            // lblWord
            // 
            this.lblWord.Location = new System.Drawing.Point(10, 10);
            this.lblWord.Name = "lblWord";
            this.lblWord.Size = new System.Drawing.Size(100, 23);
            this.lblWord.TabIndex = 0;
            this.lblWord.Text = "Word / Phrase";
            // 
            // lblReplacement
            // 
            this.lblReplacement.Location = new System.Drawing.Point(10, 40);
            this.lblReplacement.Name = "lblReplacement";
            this.lblReplacement.Size = new System.Drawing.Size(100, 23);
            this.lblReplacement.TabIndex = 2;
            this.lblReplacement.Text = "Replacement";
            // 
            // txtWord
            // 
            this.txtWord.Location = new System.Drawing.Point(120, 8);
            this.txtWord.Name = "txtWord";
            this.txtWord.Size = new System.Drawing.Size(240, 20);
            this.txtWord.TabIndex = 1;
            // 
            // txtReplacement
            // 
            this.txtReplacement.Location = new System.Drawing.Point(120, 38);
            this.txtReplacement.Name = "txtReplacement";
            this.txtReplacement.Size = new System.Drawing.Size(240, 20);
            this.txtReplacement.TabIndex = 3;
            // 
            // btnReplace
            // 
            this.btnReplace.Location = new System.Drawing.Point(0, 0);
            this.btnReplace.Name = "btnReplace";
            this.btnReplace.Size = new System.Drawing.Size(75, 23);
            this.btnReplace.TabIndex = 0;
            this.btnReplace.Text = "Replace";
            // 
            // btnReplaceAll
            // 
            this.btnReplaceAll.Location = new System.Drawing.Point(0, 0);
            this.btnReplaceAll.Name = "btnReplaceAll";
            this.btnReplaceAll.Size = new System.Drawing.Size(75, 23);
            this.btnReplaceAll.TabIndex = 0;
            this.btnReplaceAll.Text = "Replace All";
            // 
            // btnIgnore
            // 
            this.btnIgnore.Location = new System.Drawing.Point(0, 0);
            this.btnIgnore.Name = "btnIgnore";
            this.btnIgnore.Size = new System.Drawing.Size(75, 23);
            this.btnIgnore.TabIndex = 0;
            this.btnIgnore.Text = "Ignore";
            // 
            // btnIgnoreAll
            // 
            this.btnIgnoreAll.Location = new System.Drawing.Point(0, 0);
            this.btnIgnoreAll.Name = "btnIgnoreAll";
            this.btnIgnoreAll.Size = new System.Drawing.Size(75, 23);
            this.btnIgnoreAll.TabIndex = 0;
            this.btnIgnoreAll.Text = "Ignore All";
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(0, 0);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 0;
            this.btnCancel.Text = "Cancel";
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(0, 0);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "Close";
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(61, 4);
            // 
            // mainLayout
            // 
            this.mainLayout.ColumnCount = 1;
            this.mainLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.mainLayout.Controls.Add(this.textBoxInput, 0, 0);
            this.mainLayout.Controls.Add(this.tabControlModes, 0, 1);
            this.mainLayout.Controls.Add(this.bottomPanel, 0, 2);
            this.mainLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainLayout.Location = new System.Drawing.Point(0, 0);
            this.mainLayout.Name = "mainLayout";
            this.mainLayout.RowCount = 3;
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.mainLayout.Size = new System.Drawing.Size(150, 150);
            this.mainLayout.TabIndex = 1;
            // 
            // bottomPanel
            // 
            this.bottomPanel.Controls.Add(this.lblWord);
            this.bottomPanel.Controls.Add(this.txtWord);
            this.bottomPanel.Controls.Add(this.lblReplacement);
            this.bottomPanel.Controls.Add(this.txtReplacement);
            this.bottomPanel.Controls.Add(this.buttonPanel);
            this.bottomPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.bottomPanel.Location = new System.Drawing.Point(3, 27);
            this.bottomPanel.Name = "bottomPanel";
            this.bottomPanel.Padding = new System.Windows.Forms.Padding(10);
            this.bottomPanel.Size = new System.Drawing.Size(144, 120);
            this.bottomPanel.TabIndex = 2;
            // 
            // buttonPanel
            // 
            this.buttonPanel.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.buttonPanel.Location = new System.Drawing.Point(10, 70);
            this.buttonPanel.Name = "buttonPanel";
            this.buttonPanel.Size = new System.Drawing.Size(124, 40);
            this.buttonPanel.TabIndex = 4;
            this.buttonPanel.WrapContents = false;
            // 
            // SuggestionPaneControl
            // 
            this.Controls.Add(this.mainLayout);
            this.Name = "SuggestionPaneControl";
            this.tabControlModes.ResumeLayout(false);
            this.tabPageAbbreviation.ResumeLayout(false);
            this.tabPageReverse.ResumeLayout(false);
            this.tabPageDictionary.ResumeLayout(false);
            this.mainLayout.ResumeLayout(false);
            this.mainLayout.PerformLayout();
            this.bottomPanel.ResumeLayout(false);
            this.bottomPanel.PerformLayout();
            this.ResumeLayout(false);

        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            ApplyButtonStyle(btnReplace);
            ApplyButtonStyle(btnReplaceAll);
            ApplyButtonStyle(btnIgnore);
            ApplyButtonStyle(btnIgnoreAll);
            ApplyButtonStyle(btnCancel);
            ApplyButtonStyle(btnClose);

            buttonPanel.Controls.Add(btnReplace);
            buttonPanel.Controls.Add(btnReplaceAll);
            buttonPanel.Controls.Add(btnIgnore);
            buttonPanel.Controls.Add(btnIgnoreAll);
            buttonPanel.Controls.Add(btnCancel);
            buttonPanel.Controls.Add(btnClose);

            // Events (safe here)
            textBoxInput.TextChanged += textBoxInput_TextChanged;
            btnReplace.Click += btnReplace_Click;
            btnReplaceAll.Click += btnReplaceAll_Click;
            btnIgnore.Click += btnIgnore_Click;
            btnIgnoreAll.Click += btnIgnoreAll_Click;
            btnCancel.Click += btnCancel_Click;
            btnClose.Click += btnClose_Click;
        }


        // ===== MODERN BUTTON STYLE =====
        private void ApplyButtonStyle(Button btn)
        {
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
            btn.BackColor = Color.FromArgb(240, 242, 245);
            btn.ForeColor = Color.Black;
            btn.Height = 30;
            btn.Width = 140;
            btn.Margin = new Padding(5, 0, 0, 0);
        }


        // ===== KEYBOARD SHORTCUTS =====
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
            else if (e.KeyCode == Keys.Escape)
            {
                btnCancel.PerformClick();
                e.Handled = true;
            }
        }
    }
}
