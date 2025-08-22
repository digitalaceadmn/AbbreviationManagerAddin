using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using static AbbreviationWordAddin.ThisAddIn;

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

            tabControlModes.SelectedIndexChanged += TabControlModes_SelectedIndexChanged;

            // --- Setup Tab 1 (Abbreviations) ---
            this.textBoxInput.TextChanged += (s, e) =>
            {
                if (tabControlModes.SelectedTab == tabPageAbbreviation)
                    OnTextChanged?.Invoke(textBoxInput.Text);
            };
            this.listViewAbbrev.View = System.Windows.Forms.View.Details;
            this.listViewAbbrev.FullRowSelect = true;
            this.listViewAbbrev.Columns.Add("Word/Phrase", 120);
            this.listViewAbbrev.Columns.Add("Replacement", 200);
            this.listViewAbbrev.DoubleClick += ListView_DoubleClick;
            this.listViewAbbrev.MouseEnter += (s, e) => isSuggestionListFrozen = true;
            this.listViewAbbrev.MouseLeave += (s, e) => isSuggestionListFrozen = false;

            // --- Setup Tab 2 (Reverse Abbreviations) ---
            this.textBoxInput.TextChanged += (s, e) =>
            {
                if (tabControlModes.SelectedTab == tabPageReverse)
                    OnTextChanged?.Invoke(textBoxInput.Text);
            };
            this.listViewReverse.View = System.Windows.Forms.View.Details;
            this.listViewReverse.FullRowSelect = true;
            this.listViewReverse.Columns.Add("Replacement", 200);
            this.listViewReverse.Columns.Add("Word/Phrase", 120);
            this.listViewReverse.DoubleClick += ListView_DoubleClick;
            this.listViewReverse.MouseEnter += (s, e) => isSuggestionListFrozen = true;
            this.listViewReverse.MouseLeave += (s, e) => isSuggestionListFrozen = false;

            // --- Setup Tab 3 (Dictionary) ---
            this.listViewDictionary.View = System.Windows.Forms.View.Details;
            this.listViewDictionary.FullRowSelect = true;
            this.listViewDictionary.Columns.Add("Word/Phrase", 320);
            this.listViewDictionary.Columns.Add("Replacement", 200);
        }

        // --- Tab handling ---
        public enum Mode
        {
            Abbreviation,
            Reverse,
            Dictionary
        }

        public Mode CurrentMode
        {
            get
            {
                if (tabControlModes.SelectedTab == tabPageAbbreviation)
                    return Mode.Abbreviation;
                if (tabControlModes.SelectedTab == tabPageReverse)
                    return Mode.Reverse;
                return Mode.Dictionary;
            }
        }

        // --- Double click suggestion accept ---
        private void ListView_DoubleClick(object sender, EventArgs e)
        {
            var lv = sender as ListView;
            if (lv == null || lv.SelectedItems.Count == 0) return;

            var selected = lv.SelectedItems[0];
            string shortForm, fullForm;

            if (CurrentMode == Mode.Reverse)
            {
                fullForm = selected.SubItems[0].Text;   // full form
                shortForm = selected.SubItems[1].Text;  // abbreviation
            }
            else
            {
                shortForm = selected.SubItems[0].Text;  // abbreviation
                fullForm = selected.SubItems[1].Text;   // full form
            }

            OnSuggestionAccepted?.Invoke(shortForm, fullForm);
        }

        // --- Show suggestions in the right tab ---
        private List<(string Word, string Replacement)> lastSuggestions =
     new List<(string Word, string Replacement)>();

        public void ShowSuggestions(List<(string Word, string Replacement)> suggestions)
        {
            if (isSuggestionListFrozen)
            {
                Debug.WriteLine("Skipping suggestion refresh because list is frozen.");
                return;
            }

            // ✅ Skip refresh if suggestions are identical
            if (lastSuggestions.Count == suggestions.Count &&
                !lastSuggestions.Except(suggestions).Any())
            {
                return;
            }

            lastSuggestions = suggestions;

            if (CurrentMode == Mode.Abbreviation)
            {
                listViewAbbrev.BeginUpdate(); // avoid flicker
                listViewAbbrev.Items.Clear();
                foreach (var suggestion in suggestions)
                {
                    var item = new ListViewItem(suggestion.Word);
                    item.SubItems.Add(suggestion.Replacement);
                    listViewAbbrev.Items.Add(item);
                }
                listViewAbbrev.EndUpdate();
            }
            else if (CurrentMode == Mode.Reverse)
            {
                listViewReverse.BeginUpdate(); // avoid flicker
                listViewReverse.Items.Clear();
                foreach (var suggestion in suggestions)
                {
                    var item = new ListViewItem(suggestion.Replacement);
                    item.SubItems.Add(suggestion.Word);
                    listViewReverse.Items.Add(item);
                }
                listViewReverse.EndUpdate();
            }
        }


        // --- Set input text depending on tab ---
        public void SetInputText(string text)
        {
            if (CurrentMode == Mode.Abbreviation)
                textBoxInput.Text = text;
            else if (CurrentMode == Mode.Reverse)
                textBoxInput.Text = text;
        }

        // --- Load dictionary from Excel ---
        public void LoadDictionary(List<(string Abbrev, string FullForm)> entries)
        {

            listViewDictionary.Items.Clear();
            foreach (var entry in entries)
            {
                var item = new ListViewItem(entry.Abbrev);
                item.SubItems.Add(entry.FullForm);
                listViewDictionary.Items.Add(item);
            }
        }

        private void TabControlModes_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CurrentMode == Mode.Dictionary)
            {
                // Load abbreviations into dictionary view whenever user switches to the Dictionary tab
                var entries = new List<(string Abbrev, string FullForm)>();

                foreach (var abbreviation in AbbreviationManager.GetAllPhrases())
                {
                    try
                    {
                        var fullForm = AbbreviationManager.GetAbbreviation(abbreviation);

                        if (!string.IsNullOrEmpty(fullForm))
                            entries.Add((abbreviation, fullForm));
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        continue;
                    }
                }

                // ✅ Sort alphabetically by abbreviation
                entries = entries
                    .OrderBy(entry => entry.Abbrev, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                LoadDictionary(entries);
            }
        }

    
        private void SuggestionPaneControl_Load(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        public void LoadMatches(List<MatchResult> matches)
        {
            listViewAbbrev.Items.Clear();

            foreach (var m in matches)
            {
                var item = new ListViewItem(new string[] { m.Phrase, m.Replacement });
                item.Tag = m;
                listViewAbbrev.Items.Add(item);

                // Fixed MessageBox
                MessageBox.Show("Showing in list: " + m.Phrase, "Match Added");
            }

            if (listViewAbbrev.Items.Count > 0)
            {
                listViewAbbrev.Items[0].Selected = true;
                UpdateTextBoxes((MatchResult)listViewAbbrev.Items[0].Tag);
            }
        }


        private void UpdateTextBoxes(MatchResult match)
        {
            txtWord.Text = match.Phrase;
            txtReplacement.Text = match.Replacement;
        }

        private void listViewAbbrev_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listViewAbbrev.SelectedItems.Count > 0)
            {
                var match = (MatchResult)listViewAbbrev.SelectedItems[0].Tag;
                UpdateTextBoxes(match);
            }
        }

        private void btnReplace_Click(object sender, EventArgs e)
        {
            if (listViewAbbrev.SelectedItems.Count > 0)
            {
                var match = (MatchResult)listViewAbbrev.SelectedItems[0].Tag;

                // take latest replacement from textbox (if user edits)
                string replacement = txtReplacement.Text.Trim();
                if (string.IsNullOrEmpty(replacement))
                    replacement = match.Replacement;

                // replace in Word document (real-time)
                Globals.ThisAddIn.ReplaceAbbreviation(match.Phrase, replacement, false);

                // update UI: remove replaced item
                listViewAbbrev.Items.Remove(listViewAbbrev.SelectedItems[0]);

                // move to next item automatically
                if (listViewAbbrev.Items.Count > 0)
                {
                    listViewAbbrev.Items[0].Selected = true;
                    UpdateTextBoxes((MatchResult)listViewAbbrev.Items[0].Tag);
                }
                else
                {
                    txtWord.Text = "";
                    txtReplacement.Text = "";
                }
            }
        }

        private void btnReplaceAll_Click(object sender, EventArgs e)
        {
            string word = txtWord.Text.Trim();
            string replacement = txtReplacement.Text.Trim();

            if (!string.IsNullOrEmpty(word) && !string.IsNullOrEmpty(replacement))
                Globals.ThisAddIn.ReplaceAllDirectAbbreviations_Fast();
        }


        private void btnIgnore_Click(object sender, EventArgs e)
        {
            if (listViewAbbrev.SelectedItems.Count > 0)
            {
                // remove current item without replacing
                listViewAbbrev.Items.Remove(listViewAbbrev.SelectedItems[0]);

                // move to next automatically
                if (listViewAbbrev.Items.Count > 0)
                {
                    listViewAbbrev.Items[0].Selected = true;
                    UpdateTextBoxes((MatchResult)listViewAbbrev.Items[0].Tag);
                }
                else
                {
                    txtWord.Text = "";
                    txtReplacement.Text = "";
                }
            }
        }

        private void btnIgnoreAll_Click(object sender, EventArgs e)
        {
            // clear all matches
            listViewAbbrev.Items.Clear();
            txtWord.Text = "";
            txtReplacement.Text = "";
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            // behaves same as ignore all
            listViewAbbrev.Items.Clear();
            txtWord.Text = "";
            txtReplacement.Text = "";
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            // clear all matches and close the form
            listViewAbbrev.Items.Clear();
            txtWord.Text = "";
            txtReplacement.Text = "";
        }
    }
}
