﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿using System;
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
using Microsoft.Office.Tools;


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
        private int maxPhraseLength = 12;
        private bool isReplacing = false;

        private string lastReplacedShortForm = "";
        private string lastReplacedFullForm = "";
        private bool justUndone = false;


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

                this.Application.DocumentOpen += Application_DocumentOpen;
                ((Word.ApplicationEvents4_Event)this.Application).NewDocument += Application_NewDocument;
                //this.Application.WindowActivate += Application_WindowActivate;


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

        private void Application_DocumentOpen(Word.Document Doc)
        {
            EnsureTaskPaneVisible();
        }

        private void Application_NewDocument(Word.Document Doc)
        {
            
            EnsureTaskPaneVisible();
        }

        //private void Application_WindowActivate(Word.Document Doc, Word.Window Wn)
        //{
        //    System.Windows.Forms.MessageBox.Show(
        //                       "Taskpane3",
        //                       "AutoCorrect Entries",
        //                       System.Windows.Forms.MessageBoxButtons.OK,
        //                       System.Windows.Forms.MessageBoxIcon.Information
        //                   );
        //    EnsureTaskPaneVisible();
        //}

        private void EnsureTaskPaneVisible()
        {
            // Check if the pane already exists for this document
            foreach (CustomTaskPane pane in this.CustomTaskPanes)
            {
                if (pane.Control is SuggestionPaneControl && pane.Title == "Abbreviation Suggestions")
                {
                    pane.Visible = true;
                    return; // Already exists, just show it
                }
            }

            // If not found, create a new one
            suggestionPaneControl = new SuggestionPaneControl();
            suggestionPaneControl.OnTextChanged += SuggestionPaneControl_OnTextChanged;
            suggestionPaneControl.OnSuggestionAccepted += SuggestionPaneControl_OnSuggestionAccepted;

            suggestionTaskPane = this.CustomTaskPanes.Add(suggestionPaneControl, "Abbreviation Suggestions");
            suggestionTaskPane.Visible = true;
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
                .Where(p =>
                    p.StartsWith(inputText, StringComparison.InvariantCultureIgnoreCase) ||
                    string.Equals(p, inputText, StringComparison.InvariantCultureIgnoreCase)
                )
                .ToList();

            suggestionPaneControl.ShowSuggestions(matches);
        }

        /// <summary>
        /// Event: User accepted a suggestion.
        /// </summary>
        private void SuggestionPaneControl_OnSuggestionAccepted(string inputText, string abbreviation)
        {
            isReplacing = true;

            try
            {
                Word.Selection sel = this.Application.Selection;
                if (sel == null || sel.Range == null) return;

                string fullForm = AbbreviationManager.GetAbbreviation(abbreviation);
                if (string.IsNullOrEmpty(fullForm)) return;

                Word.Range replaceRange = sel.Range.Duplicate;

                int wordCount = inputText.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Length;
                replaceRange.MoveStart(Word.WdUnits.wdWord, -wordCount);

                replaceRange.Text = fullForm + " ";
                sel.SetRange(replaceRange.End, replaceRange.End);

                var autoCorrect = this.Application.AutoCorrect;
                if (!autoCorrect.Entries.Cast<Word.AutoCorrectEntry>().Any(entry => entry.Name == inputText))
                {
                    autoCorrect.Entries.Add(inputText, fullForm);
                }

                this.Application.ActiveWindow.SetFocus();
            }
            finally
            {
                isReplacing = false;
            }
        }


        /// <summary>
        /// Timer: Checks current word every interval.
        /// </summary>
        //private void TypingTimer_Tick(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        Word.Selection sel = this.Application.Selection;

        //        if (sel?.Range != null)
        //        {
        //            string matchedPhrase = null;
        //            string fullForm = null;

        //            for (int wordCount = maxPhraseLength; wordCount >= 1; wordCount--)
        //            {
        //                Word.Range testRange = sel.Range.Duplicate;
        //                testRange.MoveStart(Word.WdUnits.wdWord, -wordCount);
        //                testRange.MoveEnd(Word.WdUnits.wdWord, 0);

        //                string candidate = testRange.Text?.Trim();

        //                if (!string.IsNullOrEmpty(candidate))
        //                {
        //                    var matches = AbbreviationManager.GetAllPhrases()
        //                        .Where(p => p.StartsWith(candidate, StringComparison.InvariantCultureIgnoreCase))
        //                        .ToList();

        //                    if (matches.Count > 0)
        //                    {
        //                        var exactMatch = matches
        //                            .FirstOrDefault(p => p.Equals(candidate, StringComparison.InvariantCultureIgnoreCase));

        //                        if (!string.IsNullOrEmpty(exactMatch))
        //                        {
        //                            // Are there longer possible matches?
        //                            bool hasLongerMatch = matches.Any(p =>
        //                            {
        //                                var words = p.Split(' ');
        //                                var candidateWords = candidate.Split(' ');
        //                                return words.Length > candidateWords.Length;
        //                            });

        //                            if (hasLongerMatch)
        //                            {
        //                                string nextWord = GetNextWord(sel);

        //                                bool nextWordMatchesLonger = matches.Any(p =>
        //                                {
        //                                    var words = p.Split(' ');
        //                                    var candidateWords = candidate.Split(' ');

        //                                    if (words.Length > candidateWords.Length)
        //                                    {
        //                                        string expectedNextWord = words[candidateWords.Length];
        //                                        return expectedNextWord.StartsWith(nextWord, StringComparison.InvariantCultureIgnoreCase);
        //                                    }
        //                                    return false;
        //                                });

        //                                if (!string.IsNullOrWhiteSpace(nextWord))
        //                                {
        //                                    if (!nextWordMatchesLonger)
        //                                    {
        //                                        ReplaceWithFullForm(exactMatch, testRange, sel);
        //                                        return;
        //                                    }
        //                                    else
        //                                    {
        //                                        this.lastWord = candidate;
        //                                        suggestionPaneControl.SetInputText(candidate);
        //                                        suggestionPaneControl.ShowSuggestions(matches);
        //                                        return;
        //                                    }
        //                                }
        //                                else
        //                                {
        //                                    // Next word not typed yet → just wait
        //                                    this.lastWord = candidate;
        //                                    suggestionPaneControl.SetInputText(candidate);
        //                                    suggestionPaneControl.ShowSuggestions(matches);
        //                                    return;
        //                                }
        //                            }
        //                            else
        //                            {
        //                                // No longer match possible → expand if user typed space or next word that does not extend
        //                                string nextWord = GetNextWord(sel);
        //                                if (!string.IsNullOrWhiteSpace(nextWord))
        //                                {
        //                                    ReplaceWithFullForm(exactMatch, testRange, sel);
        //                                    return;
        //                                }
        //                                else
        //                                {
        //                                    if (sel.Range.Start > 0)
        //                                    {
        //                                        Word.Range charRange = sel.Range.Duplicate;
        //                                        charRange.MoveStart(Word.WdUnits.wdCharacter, -1);
        //                                        string lastChar = charRange.Text;

        //                                        if (lastChar == " ")
        //                                        {
        //                                            ReplaceWithFullForm(exactMatch, testRange, sel);
        //                                            return;
        //                                        }
        //                                    }
        //                                }
        //                            }
        //                        }
        //                        else
        //                        {
        //                            // Not exact yet → keep showing suggestions
        //                            this.lastWord = candidate;
        //                            suggestionPaneControl.SetInputText(candidate);
        //                            suggestionPaneControl.ShowSuggestions(matches);
        //                            return;
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Debug.WriteLine("Error in TypingTimer_Tick: " + ex.Message);
        //    }
        //}

        //private void TypingTimer_Tick(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        Word.Selection sel = this.Application.Selection;

        //        if (sel?.Range != null)
        //        {
        //            string matchedPhrase = null;
        //            string fullForm = null;

        //            for (int wordCount = maxPhraseLength; wordCount >= 1; wordCount--)
        //            {
        //                Word.Range testRange = sel.Range.Duplicate;
        //                testRange.MoveStart(Word.WdUnits.wdWord, -wordCount);
        //                testRange.MoveEnd(Word.WdUnits.wdWord, 0);

        //                string candidate = testRange.Text?.Trim();

        //                if (!string.IsNullOrEmpty(candidate))
        //                {
        //                    var matches = AbbreviationManager.GetAllPhrases()
        //                        .Where(p => p.StartsWith(candidate, StringComparison.InvariantCultureIgnoreCase))
        //                        .ToList();

        //                    if (matches.Count > 0)
        //                    {
        //                        var exactMatch = matches
        //                            .FirstOrDefault(p => p.Equals(candidate, StringComparison.InvariantCultureIgnoreCase));

        //                        if (!string.IsNullOrEmpty(exactMatch))
        //                        {
        //                            // Check if there is a longer possible match
        //                            bool hasLongerMatch = matches.Any(p =>
        //                            {
        //                                var words = p.Split(' ');
        //                                var candidateWords = candidate.Split(' ');
        //                                return words.Length > candidateWords.Length;
        //                            });

        //                            if (hasLongerMatch)
        //                            {
        //                                // Wait for next word
        //                                string nextWord = GetNextWord(sel);
        //                                bool nextWordMatchesLonger = matches.Any(p =>
        //                                {
        //                                    var words = p.Split(' ');
        //                                    var candidateWords = candidate.Split(' ');

        //                                    if (words.Length > candidateWords.Length)
        //                                    {
        //                                        string expectedNextWord = words[candidateWords.Length];
        //                                        return expectedNextWord.StartsWith(nextWord, StringComparison.InvariantCultureIgnoreCase);
        //                                    }
        //                                    return false;
        //                                });

        //                                if (!string.IsNullOrWhiteSpace(nextWord) && !nextWordMatchesLonger)
        //                                {
        //                                    // Next word does not match → expand shorter match
        //                                    ReplaceWithFullForm(exactMatch, testRange, sel);
        //                                    return;
        //                                }
        //                                else
        //                                {
        //                                    // Wait: show suggestions for longer options
        //                                    this.lastWord = candidate;
        //                                    suggestionPaneControl.SetInputText(candidate);
        //                                    suggestionPaneControl.ShowSuggestions(matches);
        //                                    return;
        //                                }
        //                            }
        //                            else
        //                            {
        //                                // ✅ FIX: No longer match possible
        //                                string nextWord = GetNextWord(sel);
        //                                if (!string.IsNullOrWhiteSpace(nextWord))
        //                                {
        //                                    // There is a next word → expand now
        //                                    ReplaceWithFullForm(exactMatch, testRange, sel);
        //                                    return;
        //                                }
        //                                else
        //                                {
        //                                    // No next word → expand only if space typed
        //                                    if (sel.Range.Start > 0)
        //                                    {
        //                                        Word.Range charRange = sel.Range.Duplicate;
        //                                        charRange.MoveStart(Word.WdUnits.wdCharacter, -1);
        //                                        string lastChar = charRange.Text;

        //                                        if (lastChar == " ")
        //                                        {
        //                                            ReplaceWithFullForm(exactMatch, testRange, sel);
        //                                            return;
        //                                        }
        //                                    }
        //                                }
        //                            }
        //                        }
        //                        else
        //                        {
        //                            // Not exact yet → keep showing suggestions
        //                            this.lastWord = candidate;
        //                            suggestionPaneControl.SetInputText(candidate);
        //                            suggestionPaneControl.ShowSuggestions(matches);
        //                            return;
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Debug.WriteLine("Error in TypingTimer_Tick: " + ex.Message);
        //    }
        //}

        private void TypingTimer_Tick(object sender, EventArgs e)
        {
            if (isReplacing) return;
            try
            {
                Word.Selection sel = this.Application.Selection;
                if (sel?.Range == null) return;

                for (int wordCount = maxPhraseLength; wordCount >= 1; wordCount--)
                {
                    Word.Range testRange = sel.Range.Duplicate;
                    testRange.MoveStart(Word.WdUnits.wdWord, -wordCount);
                    testRange.MoveEnd(Word.WdUnits.wdWord, 0);

                    string candidate = testRange.Text?.Trim();
                    if (string.IsNullOrEmpty(candidate)) continue;

                    if (!string.IsNullOrEmpty(lastReplacedShortForm) && !string.IsNullOrEmpty(lastReplacedFullForm))
                    {
                        if (string.Equals(candidate, lastReplacedShortForm, StringComparison.InvariantCultureIgnoreCase)
                            && !string.Equals(candidate, lastReplacedFullForm, StringComparison.InvariantCultureIgnoreCase))
                        {
                            justUndone = true;
                            Debug.WriteLine($"Detected undo for: {lastReplacedShortForm}");
                        }
                    }

                    var matches = AbbreviationManager.GetAllPhrases()
                        .Where(p => p.StartsWith(candidate, StringComparison.InvariantCultureIgnoreCase))
                        .ToList();

                    if (matches.Count == 0) continue;

                    bool hasExact = matches.Any(p =>
                        string.Equals(p, candidate, StringComparison.InvariantCultureIgnoreCase));

                    bool hasLonger = matches.Any(p =>
                        p.Split(' ').Length > candidate.Split(' ').Length);

                    if (hasExact && !hasLonger)
                    {
                        if (justUndone && string.Equals(candidate, lastReplacedShortForm, StringComparison.InvariantCultureIgnoreCase))
                        {
                            Debug.WriteLine($"Skipping replacement for {candidate} because it was just undone.");
                            return;
                        }

                        // Only exact, no longer — check for trailing space
                        if (IsLastCharSpace(sel))
                        {
                            ReplaceWithFullForm(candidate, testRange, sel);

                            lastReplacedShortForm = candidate;
                            lastReplacedFullForm = GetFullFormFor(candidate); 

                            justUndone = false; 
                        }
                        return;
                    }
                    else if (hasExact && hasLonger)
                    {
                        // Wait — possible longer match, show suggestions
                        suggestionPaneControl.SetInputText(candidate);
                        suggestionPaneControl.ShowSuggestions(matches);
                        return;
                    }
                    else
                    {
                        // No exact, but possible matches — show suggestions
                        suggestionPaneControl.SetInputText(candidate);
                        suggestionPaneControl.ShowSuggestions(matches);
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Error in TypingTimer_Tick: " + ex.Message);
            }
        }

        private bool IsLastCharSpace(Word.Selection sel)
        {
            if (sel.Range.Start > 0)
            {
                Word.Range charRange = sel.Range.Duplicate;
                charRange.MoveStart(Word.WdUnits.wdCharacter, -1);
                string lastChar = charRange.Text;
                return lastChar == " ";
            }
            return false;
        }


        private string GetFullFormFor(string shortForm)
        {
            // Example: adjust to your lookup logic
            return AbbreviationManager.GetAbbreviation(shortForm);
        }


        private void ReplaceWithFullForm(string matchedPhrase, Word.Range replaceRange, Word.Selection sel)
        {
            string fullForm = AbbreviationManager.GetAbbreviation(matchedPhrase);
            if (!string.IsNullOrEmpty(fullForm))
            {
                replaceRange.Text = fullForm + " ";
                sel.SetRange(replaceRange.End, replaceRange.End);
                this.Application.AutoCorrect.Entries.Add(matchedPhrase, fullForm);
                this.lastWord = matchedPhrase;
            }
        }

        private string GetNextWord(Word.Selection sel)
        {
            Word.Range testRange = sel.Range.Duplicate;
            testRange.MoveStart(Word.WdUnits.wdWord, -1);
            testRange.MoveEnd(Word.WdUnits.wdWord, 1);

            string[] words = testRange.Text.Trim().Split(' ');
            if (words.Length > 1)
            {
                return words.Last();
            }

            return "";
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
