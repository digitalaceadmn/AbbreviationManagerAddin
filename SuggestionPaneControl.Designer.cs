namespace AbbreviationWordAddin
{
    partial class SuggestionPaneControl
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.TextBox textBoxInput;
        private System.Windows.Forms.TabControl tabControlModes;
        private System.Windows.Forms.TabPage tabPageAbbreviation;
        private System.Windows.Forms.ListView listViewAbbrev;
        private System.Windows.Forms.TabPage tabPageReverse;
        private System.Windows.Forms.ListView listViewReverse;
        private System.Windows.Forms.TabPage tabPageDictionary;
        private System.Windows.Forms.ListView listViewDictionary;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.textBoxInput = new System.Windows.Forms.TextBox();
            this.tabControlModes = new System.Windows.Forms.TabControl();
            this.tabPageAbbreviation = new System.Windows.Forms.TabPage();
            this.listViewAbbrev = new System.Windows.Forms.ListView();
            this.tabPageReverse = new System.Windows.Forms.TabPage();
            this.listViewReverse = new System.Windows.Forms.ListView();
            this.tabPageDictionary = new System.Windows.Forms.TabPage();
            this.listViewDictionary = new System.Windows.Forms.ListView();

            this.tabControlModes.SuspendLayout();
            this.tabPageAbbreviation.SuspendLayout();
            this.tabPageReverse.SuspendLayout();
            this.tabPageDictionary.SuspendLayout();
            this.SuspendLayout();

            // textBoxInput
            this.textBoxInput.Location = new System.Drawing.Point(10, 10);
            this.textBoxInput.Name = "textBoxInput";
            this.textBoxInput.Size = new System.Drawing.Size(380, 22);
            this.textBoxInput.TabIndex = 0;

            // tabControlModes
            this.tabControlModes.Controls.Add(this.tabPageAbbreviation);
            this.tabControlModes.Controls.Add(this.tabPageReverse);
            this.tabControlModes.Controls.Add(this.tabPageDictionary);
            this.tabControlModes.Location = new System.Drawing.Point(10, 40);
            this.tabControlModes.Name = "tabControlModes";
            this.tabControlModes.SelectedIndex = 0;
            this.tabControlModes.Size = new System.Drawing.Size(644, 540);
            this.tabControlModes.TabIndex = 1;

            // tabPageAbbreviation
            this.tabPageAbbreviation.Controls.Add(this.listViewAbbrev);
            this.tabPageAbbreviation.Location = new System.Drawing.Point(4, 25);
            this.tabPageAbbreviation.Name = "tabPageAbbreviation";
            this.tabPageAbbreviation.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageAbbreviation.Size = new System.Drawing.Size(636, 511);
            this.tabPageAbbreviation.Text = "Abbreviations";
            this.tabPageAbbreviation.UseVisualStyleBackColor = true;

            // listViewAbbrev
            this.listViewAbbrev.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listViewAbbrev.FullRowSelect = true;
            this.listViewAbbrev.HideSelection = false;
            this.listViewAbbrev.Location = new System.Drawing.Point(3, 3);
            this.listViewAbbrev.Name = "listViewAbbrev";
            this.listViewAbbrev.Size = new System.Drawing.Size(630, 505);
            this.listViewAbbrev.TabIndex = 0;
            this.listViewAbbrev.UseCompatibleStateImageBehavior = false;

            // tabPageReverse
            this.tabPageReverse.Controls.Add(this.listViewReverse);
            this.tabPageReverse.Location = new System.Drawing.Point(4, 25);
            this.tabPageReverse.Name = "tabPageReverse";
            this.tabPageReverse.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageReverse.Size = new System.Drawing.Size(636, 511);
            this.tabPageReverse.Text = "Reverse Abbreviations";
            this.tabPageReverse.UseVisualStyleBackColor = true;

            // listViewReverse
            this.listViewReverse.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listViewReverse.FullRowSelect = true;
            this.listViewReverse.HideSelection = false;
            this.listViewReverse.Location = new System.Drawing.Point(3, 3);
            this.listViewReverse.Name = "listViewReverse";
            this.listViewReverse.Size = new System.Drawing.Size(630, 505);
            this.listViewReverse.TabIndex = 0;
            this.listViewReverse.UseCompatibleStateImageBehavior = false;

            // tabPageDictionary
            this.tabPageDictionary.Controls.Add(this.listViewDictionary);
            this.tabPageDictionary.Location = new System.Drawing.Point(4, 25);
            this.tabPageDictionary.Name = "tabPageDictionary";
            this.tabPageDictionary.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageDictionary.Size = new System.Drawing.Size(636, 511);
            this.tabPageDictionary.Text = "Dictionary";
            this.tabPageDictionary.UseVisualStyleBackColor = true;

            // listViewDictionary
            this.listViewDictionary.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listViewDictionary.FullRowSelect = true;
            this.listViewDictionary.HideSelection = false;
            this.listViewDictionary.Location = new System.Drawing.Point(3, 3);
            this.listViewDictionary.Name = "listViewDictionary";
            this.listViewDictionary.Size = new System.Drawing.Size(630, 505);
            this.listViewDictionary.TabIndex = 0;
            this.listViewDictionary.UseCompatibleStateImageBehavior = false;

            // SuggestionPaneControl
            this.Controls.Add(this.tabControlModes);
            this.Controls.Add(this.textBoxInput);
            this.Name = "SuggestionPaneControl";
            this.Size = new System.Drawing.Size(700, 600);

            this.tabControlModes.ResumeLayout(false);
            this.tabPageAbbreviation.ResumeLayout(false);
            this.tabPageReverse.ResumeLayout(false);
            this.tabPageDictionary.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
