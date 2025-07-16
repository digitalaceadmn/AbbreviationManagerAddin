﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;
using Action = System.Action;
using System.Windows.Forms;


namespace AbbreviationWordAddin
{
    public partial class ThisAddIn
    {
        public bool reloadAbbrDataFromDict = false; 
        private const int CHUNK_SIZE = 1000;
        string lastLoadedVersion = Properties.Settings.Default.LastLoadedAbbreviationVersion;
        string currentVersion = Properties.Settings.Default.AbbreviationDataVersion;
        private Microsoft.Office.Tools.CustomTaskPane suggestionTaskPane;
        private SuggestionPaneControl suggestionPaneControl;

        private string lastWord = "";
        private Timer typingTimer;


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            try
            {
               
                System.Windows.Forms.MessageBox.Show(
                    "lastLoadedVersion" + lastLoadedVersion + "currentVersion" + currentVersion,
                    "Abbreviation Loading status",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information
                 );

                var autoCorrect = Globals.ThisAddIn.Application.AutoCorrect;
                for (int i = autoCorrect.Entries.Count; i >= 1; i--)
                {
                    autoCorrect.Entries[i].Delete();
                }


                if (lastLoadedVersion != currentVersion)
                {
                    // Version changed → clear file cache
                    System.Windows.Forms.MessageBox.Show(
                        "Clear cache because lastLoadedVersion" + lastLoadedVersion + "currentVersion" + currentVersion,
                        "Abbreviation Loading status",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information
                     );
                    reloadAbbrDataFromDict = true;
                    AbbreviationManager.ClearCacheFile();
                    Properties.Settings.Default.IsAutoCorrectLoaded = false;
                    Properties.Settings.Default.LastLoadedAbbreviationVersion = currentVersion;
                    Properties.Settings.Default.Save();
                    Properties.Settings.Default.Reload();
                }

                AbbreviationManager.LoadAbbreviations(); 

                loadAllAbbreviaitons();

                suggestionPaneControl = new SuggestionPaneControl();
                suggestionPaneControl.OnTextChanged += SuggestionPaneControl_OnTextChanged;
                suggestionPaneControl.OnSuggestionAccepted += SuggestionPaneControl_OnSuggestionAccepted;

                suggestionTaskPane = this.CustomTaskPanes.Add(suggestionPaneControl, "Abbreviation Suggestions");
                suggestionTaskPane.Visible = true;

                typingTimer = new Timer { Interval = 300 };
                typingTimer.Tick += TypingTimer_Tick;
                typingTimer.Start();


            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Error during startup: " + ex.Message,
                    "Startup Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning
                );
            }
        }

        private void loadAllAbbreviaitons()
        {
            var application = Globals.ThisAddIn.Application;
            AutoCorrect autoCorrect = application.AutoCorrect;

            if (lastLoadedVersion != currentVersion)
            {
                reloadAbbrDataFromDict = true;

            }
            else
            {
                reloadAbbrDataFromDict = autoCorrect.ReplaceText;
            }


            AbbreviationManager.InitializeAutoCorrectCache(autoCorrect);

            String loadingStatusMessage = "";

            var entries = autoCorrect.Entries;
            string entryList = "";
            foreach (var entry in entries)
            {
                var acEntry = entry as Microsoft.Office.Interop.Word.AutoCorrectEntry;
                if (acEntry != null)
                {
                    entryList += $"{acEntry.Name} => {acEntry.Value}\n";
                }
            }

            if (!Properties.Settings.Default.IsAutoCorrectLoaded)
            {
                if (reloadAbbrDataFromDict)
                {
                    System.Windows.Forms.MessageBox.Show(
                               "Loading Latest Abbreviation",
                               "AutoCorrect Entries",
                               System.Windows.Forms.MessageBoxButtons.OK,
                               System.Windows.Forms.MessageBoxIcon.Information
                           );
                    foreach (var abbreviation in AbbreviationManager.GetAllPhrases())
                    {
                        try
                        {
                            
                            string fullForm = AbbreviationManager.GetAbbreviation(abbreviation);
                            if (!string.IsNullOrEmpty(fullForm))
                            {
                                autoCorrect.ReplaceText = true;
                                //autoCorrect.Entries.Add(abbreviation, fullForm);

                                //var template = application.NormalTemplate;

                                //Word.Document tempDoc = application.Documents.Add(Visible: false);
                                //Word.Range tempRange = tempDoc.Content;
                                //tempRange.Text = fullForm;

                                //template.AutoTextEntries.Add(abbreviation, tempRange);

                                //tempDoc.Close(false);
                            }
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {
                            loadingStatusMessage += ", " + abbreviation;
                            continue;
                        }
                    }

                    if (loadingStatusMessage != "")
                    {
                        System.Windows.Forms.MessageBox.Show(
                            "Abbreviations Loaded. Below phrases were already present in the abbreviation list: " + loadingStatusMessage,
                            "Abbreviation Loading status",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Information
                        );
                    }
                }
            } 

            Properties.Settings.Default.IsAutoCorrectLoaded = true;
            Properties.Settings.Default.Save();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //AbbreviationManager.ClearAutoCorrectCache();

        }

        public void ToggleAbbreviationReplacement(bool enable)
        {
            if (!Properties.Settings.Default.IsAutoCorrectLoaded)
            {
                reloadAbbrDataFromDict = enable;
                if (reloadAbbrDataFromDict)
                {
                    AbbreviationManager.InitializeAutoCorrectCache(this.Application.AutoCorrect);
                    System.Windows.Forms.MessageBox.Show("Abbreviation Replacement Enabled", "Status");
                }
                else
                {
                    AbbreviationManager.ClearAutoCorrectCache();
                    System.Windows.Forms.MessageBox.Show("Abbreviation Replacement Disabled", "Status");
                }
            }
                
        }


        public void ReplaceAllAbbreviations()
        {
            var progressForm = new ProgressForm();

            var syncContext = System.Threading.SynchronizationContext.Current;
            bool completed = false;
            Exception processError = null;

            var progressThread = new System.Threading.Thread(() =>
            {
                try
                {
                    Word.Document doc = null;
                    syncContext.Send(_ =>
                    {
                        doc = this.Application.ActiveDocument;
                        this.Application.ScreenUpdating = false; 
                        this.Application.DisplayStatusBar = false; 
                        this.Application.Options.ReplaceSelection = false; 
                    }, null);

                    if (reloadAbbrDataFromDict)
                    {
                        syncContext.Send(_ =>
                        {
                            AbbreviationManager.InitializeAutoCorrectCache(this.Application.AutoCorrect);
                        }, null);
                    }

                    int totalWords = 0;
                    syncContext.Send(_ =>
                    {
                        totalWords = doc.Words.Count;
                    }, null);

                    int totalChunks = (totalWords + CHUNK_SIZE - 1) / CHUNK_SIZE;
                    int currentChunk = 0;


                    for (int startIndex = 1; startIndex <= totalWords && !completed; startIndex += CHUNK_SIZE)
                    {
                        currentChunk++;
                        int endIndex = Math.Min(startIndex + CHUNK_SIZE - 1, totalWords);

                        int percentage = (currentChunk * 100) / totalChunks;
                        progressForm.UpdateProgress(percentage, $"Processing chunk {currentChunk} of {totalChunks}...");

                        syncContext.Send(_ =>
                        {
                            try
                            {
                                Word.Range chunkRange = doc.Range(doc.Words[startIndex].Start, doc.Words[endIndex].End);
                                string chunkText = chunkRange.Text;
                                bool hasMatches = false;

                                foreach (var phrase in AbbreviationManager.GetAllPhrases())
                                {
                                    if (chunkText.IndexOf(phrase, StringComparison.OrdinalIgnoreCase) != -1)
                                    {
                                        hasMatches = true;
                                        break;
                                    }
                                }

                                if (hasMatches)
                                {
                                    foreach (var phrase in AbbreviationManager.GetAllPhrases())
                                    {
                                        string replacement = AbbreviationManager.GetFromAutoCorrectCache(phrase)
                                            ?? AbbreviationManager.GetAbbreviation(phrase);

                                        if (chunkText.IndexOf(phrase, StringComparison.OrdinalIgnoreCase) != -1)
                                        {
                                            var find = chunkRange.Find;
                                            find.ClearFormatting();
                                            find.Text = phrase;
                                            find.Forward = true;
                                            find.Format = false;
                                            find.MatchCase = false;
                                            find.MatchWholeWord = true;
                                            find.MatchWildcards = false;
                                            find.MatchSoundsLike = false;
                                            find.MatchAllWordForms = false;
                                            find.Wrap = Word.WdFindWrap.wdFindContinue;

                                            find.Replacement.ClearFormatting();
                                            find.Replacement.Text = replacement;

                                            find.Execute(
                                                FindText: phrase,
                                                MatchCase: false,
                                                MatchWholeWord: true,
                                                MatchWildcards: false,
                                                MatchSoundsLike: false,
                                                MatchAllWordForms: false,
                                                Forward: true,
                                                Wrap: Word.WdFindWrap.wdFindContinue,
                                                Format: false,
                                                ReplaceWith: replacement,
                                                Replace: Word.WdReplace.wdReplaceAll
                                            );
                                        }
                                    }
                                }

                                if (chunkRange != null)
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(chunkRange);
                            }
                            catch (Exception ex)
                            {
                                processError = ex;
                                completed = true; 
                            }
                        }, null);
                    }
                }
                catch (Exception ex)
                {
                    processError = ex;
                }
                finally
                {
                    syncContext.Send(_ =>
                    {
                        this.Application.ScreenUpdating = true; 
                        this.Application.DisplayStatusBar = true; 
                        this.Application.Options.ReplaceSelection = true; 
                    }, null);

                    completed = true;
                    syncContext.Post(_ => progressForm.Close(), null);
                }
            });

            progressThread.Start();
            progressForm.ShowDialog();

        }


        public void HighlightAllAbbreviations()
        {
            var progressForm = new ProgressForm();
            var syncContext = System.Threading.SynchronizationContext.Current;
            bool completed = false;
            Exception processError = null;

            var progressThread = new System.Threading.Thread(() =>
            {
                try
                {
                    Word.Document doc = null;
                    syncContext.Send(_ =>
                    {
                        doc = this.Application.ActiveDocument;
                        this.Application.ScreenUpdating = false; // Disable screen updating to prevent flickering
                        this.Application.DisplayStatusBar = false; // Disable status bar updates
                        this.Application.Options.ReplaceSelection = false; // Disable selection replacement
                    }, null);

                    if (!AbbreviationManager.IsAutoCorrectCacheInitialized())
                    {
                        syncContext.Send(_ =>
                        {
                            AbbreviationManager.InitializeAutoCorrectCache(this.Application.AutoCorrect);
                        }, null);
                    }

                    int totalWords = 0;
                    syncContext.Send(_ =>
                    {
                        totalWords = doc.Words.Count;
                    }, null);

                    int totalChunks = (totalWords + CHUNK_SIZE - 1) / CHUNK_SIZE;
                    int currentChunk = 0;

                    // Process document in chunks
                    for (int startIndex = 1; startIndex <= totalWords && !completed; startIndex += CHUNK_SIZE)
                    {
                        currentChunk++;
                        int endIndex = Math.Min(startIndex + CHUNK_SIZE - 1, totalWords);

                        // Update progress
                        int percentage = (currentChunk * 100) / totalChunks;
                        progressForm.UpdateProgress(percentage, $"Processing chunk {currentChunk} of {totalChunks}...");

                        // Process chunk on UI thread
                        syncContext.Send(_ =>
                        {
                            try
                            {
                                Word.Range chunkRange = doc.Range(doc.Words[startIndex].Start, doc.Words[endIndex].End);
                                string chunkText = chunkRange.Text;
                                bool hasMatches = false;

                                // Quick check if chunk contains any potential matches
                                foreach (var phrase in AbbreviationManager.GetAllPhrases())
                                {
                                    if (chunkText.IndexOf(phrase, StringComparison.OrdinalIgnoreCase) != -1)
                                    {
                                        hasMatches = true;
                                        break;
                                    }
                                }

                                if (hasMatches)
                                {
                                    foreach (var phrase in AbbreviationManager.GetAllPhrases())
                                    {
                                        if (chunkText.IndexOf(phrase, StringComparison.OrdinalIgnoreCase) != -1)
                                        {
                                            var find = chunkRange.Find;
                                            find.ClearFormatting();
                                            find.Text = phrase;
                                            find.Forward = true;
                                            find.Format = true;
                                            find.MatchCase = false;
                                            find.MatchWholeWord = true;
                                            find.MatchWildcards = false;
                                            find.MatchSoundsLike = false;
                                            find.MatchAllWordForms = false;
                                            find.Wrap = Word.WdFindWrap.wdFindContinue;

                                            find.Replacement.ClearFormatting();
                                            find.Replacement.Font.Color = Word.WdColor.wdColorRed;
                                            find.Replacement.Text = phrase;  // Keep the same text, just change color

                                            // Execute highlighting
                                            find.Execute(
                                                FindText: phrase,
                                                MatchCase: false,
                                                MatchWholeWord: true,
                                                MatchWildcards: false,
                                                MatchSoundsLike: false,
                                                MatchAllWordForms: false,
                                                Forward: true,
                                                Wrap: Word.WdFindWrap.wdFindContinue,
                                                Format: true,
                                                ReplaceWith: phrase,
                                                Replace: Word.WdReplace.wdReplaceAll
                                            );
                                        }
                                    }
                                }

                                // Release COM objects
                                if (chunkRange != null)
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(chunkRange);
                            }
                            catch (Exception ex)
                            {
                                processError = ex;
                                completed = true; // Stop processing on error
                            }
                        }, null);
                    }
                }
                catch (Exception ex)
                {
                    processError = ex;
                }
                finally
                {
                    syncContext.Send(_ =>
                    {
                        this.Application.ScreenUpdating = true; 
                        this.Application.DisplayStatusBar = true; 
                        this.Application.Options.ReplaceSelection = true; 
                        this.Application.Visible = true; 
                    }, null);

                    completed = true;
                    syncContext.Post(_ => progressForm.Close(), null);
                }
            });

            progressThread.Start();
            progressForm.ShowDialog();

            //System.Windows.Forms.MessageBox.Show("HighlightAllAbbreviations Method executed", "Status");

            if (processError != null)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Error during highlighting: " + processError.Message,
                    "Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error
                );
            }
        }

        /// <summary>
        /// Event: User typed text in the pane input box.
        /// </summary>
        private void SuggestionPaneControl_OnTextChanged(string inputText)
        {
            var matches = AbbreviationManager.GetAllPhrases()
                .Where(p => p.IndexOf(inputText, StringComparison.InvariantCultureIgnoreCase) >= 0)
                .ToList();

            suggestionPaneControl.ShowSuggestions(matches);
        }

        /// <summary>
        /// Event: User accepted a suggestion.
        /// </summary>
        private void SuggestionPaneControl_OnSuggestionAccepted(string abbreviation)
        {
            Word.Selection sel = this.Application.Selection;
            if (sel == null || sel.Range == null) return;

            string fullForm = AbbreviationManager.GetAbbreviation(abbreviation);

            if (string.IsNullOrEmpty(fullForm))
            {
                Debug.WriteLine($"No full form found for abbreviation '{abbreviation}'");
                return;
            }

            Word.Range wordRange = sel.Range.Duplicate;

            if (wordRange != null)
            {
                int wordCount = abbreviation.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Length;

                wordRange.MoveStart(Word.WdUnits.wdWord, -wordCount);

                wordRange.Text = fullForm + " ";

                sel.SetRange(wordRange.End, wordRange.End);
            }

            try
            {
                var autoCorrect = this.Application.AutoCorrect;
                autoCorrect.ReplaceText = true;
                autoCorrect.Entries.Add(abbreviation, fullForm);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                Debug.WriteLine($"AutoCorrect entry for '{abbreviation}' could not be added.");
            }
        }



        /// <summary>
        /// Timer: Checks current word every interval.
        /// </summary>
        private void TypingTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                Word.Selection sel = this.Application.Selection;

                if (sel != null && sel.Range != null)
                {
                    Word.Range range = sel.Range.Duplicate;

                    if (range != null)
                    {
                        // Capture last 2 words
                        range.MoveStart(Word.WdUnits.wdWord, -2);
                        string lastTwoWords = range.Text?.Trim();

                        range.MoveStart(Word.WdUnits.wdWord, 1);
                        string lastWord = range.Text?.Trim();

                        string phraseToUse = null;

                        if (!string.IsNullOrEmpty(lastTwoWords) && AbbreviationManager.GetAbbreviation(lastTwoWords) != null)
                        {
                            phraseToUse = lastTwoWords;
                        }
                        else if (!string.IsNullOrEmpty(lastWord) && AbbreviationManager.GetAbbreviation(lastWord) != null)
                        {
                            phraseToUse = lastWord;
                        }

                        if (phraseToUse != null && phraseToUse != this.lastWord)
                        {
                            this.lastWord = phraseToUse;
                            suggestionPaneControl.SetInputText(phraseToUse);
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Error in TypingTimer_Tick: " + ex.Message);
            }
        }




        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
