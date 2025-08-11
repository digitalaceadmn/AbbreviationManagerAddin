namespace AbbreviationWordAddin
{
    partial class SuggestionPaneControl
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.TextBox textBoxInput;
        private System.Windows.Forms.ListView listBoxSuggestions;

        public void SetSuggestionsVisible(bool visible)
        {
            listBoxSuggestions.Visible = visible;
        }


        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.textBoxInput = new System.Windows.Forms.TextBox();
            this.listBoxSuggestions = new System.Windows.Forms.ListView();
            this.SuspendLayout();
            // 
            // textBoxInput
            // 
            this.textBoxInput.Location = new System.Drawing.Point(10, 10);
            this.textBoxInput.Name = "textBoxInput";
            this.textBoxInput.Size = new System.Drawing.Size(380, 22);
            this.textBoxInput.TabIndex = 0;
            // 
            // listBoxSuggestions
            // 
            this.listBoxSuggestions.FullRowSelect = true;
            this.listBoxSuggestions.HideSelection = false;
            this.listBoxSuggestions.Location = new System.Drawing.Point(10, 40);
            this.listBoxSuggestions.Name = "listBoxSuggestions";
            this.listBoxSuggestions.Size = new System.Drawing.Size(580, 300);
            this.listBoxSuggestions.TabIndex = 1;
            this.listBoxSuggestions.UseCompatibleStateImageBehavior = false;
            this.listBoxSuggestions.View = System.Windows.Forms.View.Details;
            this.listBoxSuggestions.Scrollable = true;
            this.listBoxSuggestions.MultiSelect = false;

            // Add a column so scrollbars can appear
            this.listBoxSuggestions.Columns.Add("Word/Phrase", 260);

            this.listBoxSuggestions.SelectedIndexChanged += new System.EventHandler(this.listBoxSuggestions_SelectedIndexChanged);
            // 
            // SuggestionPaneControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.textBoxInput);
            this.Controls.Add(this.listBoxSuggestions);
            this.Name = "SuggestionPaneControl";
            this.Size = new System.Drawing.Size(400, 250);
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
