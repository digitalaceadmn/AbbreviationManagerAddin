using System;
using System.Windows.Forms;

namespace AbbreviationWordAddin
{
    public partial class SuggestionPaneControl : UserControl
    {
        public event Action<string> OnTextChanged;
        public event Action<string, string> OnSuggestionAccepted;

        public SuggestionPaneControl()
        {
            InitializeComponent();

            // Attach event handler:
            this.textBoxInput.TextChanged += textBoxInput_TextChanged;
            this.listBoxSuggestions.DoubleClick += listBoxSuggestions_DoubleClick;
        }

        private void textBoxInput_TextChanged(object sender, EventArgs e)
        {
            OnTextChanged?.Invoke(textBoxInput.Text);
        }

        private void listBoxSuggestions_DoubleClick(object sender, EventArgs e)
        {
            if (listBoxSuggestions.SelectedItem != null)
            {
                OnSuggestionAccepted?.Invoke(
                    textBoxInput.Text,
                    listBoxSuggestions.SelectedItem.ToString()
                );
            }
        }

        public void ShowSuggestions(System.Collections.Generic.List<string> suggestions)
        {
            listBoxSuggestions.Items.Clear();
            foreach (var item in suggestions)
            {
                listBoxSuggestions.Items.Add(item);
            }
        }

        public void SetInputText(string text)
        {
            textBoxInput.Text = text;
        }
    }
}
