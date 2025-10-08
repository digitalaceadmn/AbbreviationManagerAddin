using System;
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
            this.tabControlModes.SuspendLayout();
            this.tabPageAbbreviation.SuspendLayout();
            this.tabPageReverse.SuspendLayout();
            this.tabPageDictionary.SuspendLayout();
            this.SuspendLayout();
            // 
            // textBoxInput
            // 
            this.textBoxInput.Location = new System.Drawing.Point(10, 10);
            this.textBoxInput.Name = "textBoxInput";
            this.textBoxInput.Size = new System.Drawing.Size(380, 22);
            this.textBoxInput.TabIndex = 1;
            this.textBoxInput.TextChanged += new System.EventHandler(this.textBoxInput_TextChanged);
            // 
            // tabControlModes
            // 
            this.tabControlModes.Controls.Add(this.tabPageAbbreviation);
            this.tabControlModes.Controls.Add(this.tabPageReverse);
            this.tabControlModes.Controls.Add(this.tabPageDictionary);
            this.tabControlModes.Location = new System.Drawing.Point(10, 40);
            this.tabControlModes.Name = "tabControlModes";
            this.tabControlModes.SelectedIndex = 0;
            this.tabControlModes.Size = new System.Drawing.Size(987, 540);
            this.tabControlModes.TabIndex = 2;
            // 
            // tabPageAbbreviation
            // 
            this.tabPageAbbreviation.Controls.Add(this.listViewAbbrev);
            this.tabPageAbbreviation.Location = new System.Drawing.Point(4, 25);
            this.tabPageAbbreviation.Name = "tabPageAbbreviation";
            this.tabPageAbbreviation.Size = new System.Drawing.Size(979, 511);
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
            this.listViewAbbrev.Size = new System.Drawing.Size(979, 511);
            this.listViewAbbrev.TabIndex = 0;
            this.listViewAbbrev.UseCompatibleStateImageBehavior = false;
            // 
            // tabPageReverse
            // 
            this.tabPageReverse.Controls.Add(this.listViewReverse);
            this.tabPageReverse.Location = new System.Drawing.Point(4, 25);
            this.tabPageReverse.Name = "tabPageReverse";
            this.tabPageReverse.Size = new System.Drawing.Size(979, 511);
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
            this.listViewReverse.Size = new System.Drawing.Size(979, 511);
            this.listViewReverse.TabIndex = 0;
            this.listViewReverse.UseCompatibleStateImageBehavior = false;
            // 
            // tabPageDictionary
            // 
            this.tabPageDictionary.Controls.Add(this.listViewDictionary);
            this.tabPageDictionary.Location = new System.Drawing.Point(4, 25);
            this.tabPageDictionary.Name = "tabPageDictionary";
            this.tabPageDictionary.Size = new System.Drawing.Size(979, 511);
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
            this.listViewDictionary.Size = new System.Drawing.Size(979, 511);
            this.listViewDictionary.TabIndex = 0;
            this.listViewDictionary.UseCompatibleStateImageBehavior = false;
            // 
            // lblWord
            // 
            this.lblWord.AutoSize = true;
            this.lblWord.Location = new System.Drawing.Point(10, 600);
            this.lblWord.Name = "lblWord";
            this.lblWord.Size = new System.Drawing.Size(93, 16);
            this.lblWord.TabIndex = 3;
            this.lblWord.Text = "Word / Phrase";
            // 
            // lblReplacement
            // 
            this.lblReplacement.AutoSize = true;
            this.lblReplacement.Location = new System.Drawing.Point(10, 630);
            this.lblReplacement.Name = "lblReplacement";
            this.lblReplacement.Size = new System.Drawing.Size(88, 16);
            this.lblReplacement.TabIndex = 6;
            this.lblReplacement.Text = "Replacement";
            // 
            // txtWord
            // 
            this.txtWord.Location = new System.Drawing.Point(120, 597);
            this.txtWord.Name = "txtWord";
            this.txtWord.Size = new System.Drawing.Size(250, 22);
            this.txtWord.TabIndex = 4;
            // 
            // txtReplacement
            // 
            this.txtReplacement.Location = new System.Drawing.Point(120, 627);
            this.txtReplacement.Name = "txtReplacement";
            this.txtReplacement.Size = new System.Drawing.Size(250, 22);
            this.txtReplacement.TabIndex = 7;
            // 
            // btnReplace
            // 
            this.btnReplace.Location = new System.Drawing.Point(380, 595);
            this.btnReplace.Name = "btnReplace";
            this.btnReplace.Size = new System.Drawing.Size(75, 23);
            this.btnReplace.TabIndex = 5;
            this.btnReplace.Text = "Replace";
            this.btnReplace.Click += new System.EventHandler(this.btnReplace_Click);
            // 
            // btnReplaceAll
            // 
            this.btnReplaceAll.Location = new System.Drawing.Point(380, 625);
            this.btnReplaceAll.Name = "btnReplaceAll";
            this.btnReplaceAll.Size = new System.Drawing.Size(98, 23);
            this.btnReplaceAll.TabIndex = 8;
            this.btnReplaceAll.Text = "Replace All";
            this.btnReplaceAll.Click += new System.EventHandler(this.btnReplaceAll_Click);
            // 
            // btnIgnore
            // 
            this.btnIgnore.Location = new System.Drawing.Point(280, 660);
            this.btnIgnore.Name = "btnIgnore";
            this.btnIgnore.Size = new System.Drawing.Size(75, 23);
            this.btnIgnore.TabIndex = 11;
            this.btnIgnore.Text = "Ignore";
            this.btnIgnore.Click += new System.EventHandler(this.btnIgnore_Click);
            // 
            // btnIgnoreAll
            // 
            this.btnIgnoreAll.Location = new System.Drawing.Point(360, 660);
            this.btnIgnoreAll.Name = "btnIgnoreAll";
            this.btnIgnoreAll.Size = new System.Drawing.Size(75, 23);
            this.btnIgnoreAll.TabIndex = 12;
            this.btnIgnoreAll.Text = "Ignore All";
            this.btnIgnoreAll.Click += new System.EventHandler(this.btnIgnoreAll_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(200, 660);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 10;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(120, 660);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 9;
            this.btnClose.Text = "Close";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(61, 4);
            // 
            // SuggestionPaneControl
            // 
            this.Controls.Add(this.textBoxInput);
            this.Controls.Add(this.tabControlModes);
            this.Controls.Add(this.lblWord);
            this.Controls.Add(this.txtWord);
            this.Controls.Add(this.btnReplace);
            this.Controls.Add(this.lblReplacement);
            this.Controls.Add(this.txtReplacement);
            this.Controls.Add(this.btnReplaceAll);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnIgnore);
            this.Controls.Add(this.btnIgnoreAll);
            this.Name = "SuggestionPaneControl";
            this.Size = new System.Drawing.Size(1000, 750);
            this.tabControlModes.ResumeLayout(false);
            this.tabPageAbbreviation.ResumeLayout(false);
            this.tabPageReverse.ResumeLayout(false);
            this.tabPageDictionary.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }
    }
}
