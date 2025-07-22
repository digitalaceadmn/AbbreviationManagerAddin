namespace AbbreviationWordAddin
{
    partial class SuggestionPaneControl
    {
        private System.Windows.Forms.TextBox textBoxInput;
        private System.Windows.Forms.ListBox listBoxSuggestions;

        private void InitializeComponent()
        {
            this.textBoxInput = new System.Windows.Forms.TextBox();
            this.listBoxSuggestions = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // textBoxInput
            // 
            this.textBoxInput.Location = new System.Drawing.Point(10, 10);
            this.textBoxInput.Name = "textBoxInput";
            this.textBoxInput.Size = new System.Drawing.Size(200, 22);
            this.textBoxInput.TabIndex = 0;
            // 
            // listBoxSuggestions
            // 
            this.listBoxSuggestions.ItemHeight = 16;
            this.listBoxSuggestions.Location = new System.Drawing.Point(10, 40);
            this.listBoxSuggestions.Name = "listBoxSuggestions";
            this.listBoxSuggestions.Size = new System.Drawing.Size(350, 340);
            this.listBoxSuggestions.TabIndex = 1;
            this.listBoxSuggestions.SelectedIndexChanged += new System.EventHandler(this.listBoxSuggestions_SelectedIndexChanged);
            // 
            // SuggestionPaneControl
            // 
            this.Controls.Add(this.textBoxInput);
            this.Controls.Add(this.listBoxSuggestions);
            this.Name = "SuggestionPaneControl";
            this.Size = new System.Drawing.Size(400, 400);
            this.ResumeLayout(false);
            this.PerformLayout();

        }
    }
}
