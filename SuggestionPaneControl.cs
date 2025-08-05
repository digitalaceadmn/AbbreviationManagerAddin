using System;
using System.Collections.Generic;
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

            // TextBox handler
            this.textBoxInput.TextChanged += textBoxInput_TextChanged;

            // Setup ListView
            this.listBoxSuggestions.View = View.Details;
            this.listBoxSuggestions.FullRowSelect = true;
            this.listBoxSuggestions.Columns.Add("Word/Phrase", 200);
            this.listBoxSuggestions.Columns.Add("Replacement", 200);

            this.listBoxSuggestions.DoubleClick += listBoxSuggestions_DoubleClick;
        }

        private void textBoxInput_TextChanged(object sender, EventArgs e)
        {
            OnTextChanged?.Invoke(textBoxInput.Text);
        }

        private void listBoxSuggestions_DoubleClick(object sender, EventArgs e)
        {
            if (listBoxSuggestions.SelectedItems.Count > 0)
            {
                var selected = listBoxSuggestions.SelectedItems[0];
                string word = selected.SubItems[0].Text;
                string replacement = selected.SubItems[1].Text;

                MessageBox.Show($"You selected: {word} → {replacement}");

                OnSuggestionAccepted?.Invoke(word, replacement);
            }
        }


        public void ShowSuggestions(List<(string Word, string Replacement)> suggestions)
        {
            listBoxSuggestions.Items.Clear();
            foreach (var suggestion in suggestions)
            {
                var item = new ListViewItem(suggestion.Word);
                item.SubItems.Add(suggestion.Replacement);
                listBoxSuggestions.Items.Add(item);
            }
        }

        public void SetInputText(string text)
        {
            textBoxInput.Text = text;
        }

        private void listBoxSuggestions_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
