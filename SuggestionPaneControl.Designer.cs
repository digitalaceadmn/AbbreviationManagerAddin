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
            this.textBoxInput.Width = 200;
            // 
            // listBoxSuggestions
            // 
            this.listBoxSuggestions.Location = new System.Drawing.Point(10, 40);
            this.listBoxSuggestions.Width = 200;
            this.listBoxSuggestions.Height = 200;
            // 
            // SuggestionPaneControl
            // 
            this.Controls.Add(this.textBoxInput);
            this.Controls.Add(this.listBoxSuggestions);
            this.Name = "SuggestionPaneControl";
            this.Size = new System.Drawing.Size(220, 260);
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
