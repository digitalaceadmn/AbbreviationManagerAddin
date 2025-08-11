using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;

namespace AbbreviationWordAddin
{
    public partial class SuggestionPaneControl : UserControl
    {
        public event Action<string> OnTextChanged;
        public event Action<string, string> OnSuggestionAccepted;
        private bool isSuggestionListFrozen = false;

        public SuggestionPaneControl()
        {
            InitializeComponent();

            // TextBox handler
            this.textBoxInput.TextChanged += textBoxInput_TextChanged;

            // Setup ListView
            this.listBoxSuggestions.View = View.Details;
            this.listBoxSuggestions.FullRowSelect = true;
            this.listBoxSuggestions.Columns.Add("Replacement", 200);

            this.listBoxSuggestions.DoubleClick += listBoxSuggestions_DoubleClick;
            this.listBoxSuggestions.MouseEnter += listBoxSuggestions_MouseEnter;
            this.listBoxSuggestions.MouseLeave += listBoxSuggestions_MouseLeave;
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
                //string word = selected.SubItems[0].Text;
                //string replacement = selected.SubItems[1].Text;


                string shortForm = selected.SubItems[0].Text;  // "accounting unit"
                string fullForm = selected.SubItems[1].Text;  // "au"

                //MessageBox.Show($"You selected: {shortForm} → {fullForm}");


                OnSuggestionAccepted?.Invoke(shortForm, fullForm);

                //OnSuggestionAccepted?.Invoke(word, replacement);
            }
        }

        private void listBoxSuggestions_MouseEnter(object sender, EventArgs e)
        {
            isSuggestionListFrozen = true;
        }

        private void listBoxSuggestions_MouseLeave(object sender, EventArgs e)
        {
            isSuggestionListFrozen = false;
        }


        public void ShowSuggestions(List<(string Word, string Replacement)> suggestions)
        {
            if (isSuggestionListFrozen)
            {
                Debug.WriteLine("Skipping suggestion refresh because list is frozen.");
                return;
            }

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
