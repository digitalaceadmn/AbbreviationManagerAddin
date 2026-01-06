using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Word;
﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Action = System.Action;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;


namespace AbbreviationWordAddin
{
    public partial class ThisAddIn
    {
        public bool reloadAbbrDataFromDict = false; 
        public bool debug = false;
        private const int CHUNK_SIZE = 1000;
        string lastLoadedVersion = Properties.Settings.Default.LastLoadedAbbreviationVersion;
        string currentVersion = Properties.Settings.Default.AbbreviationDataVersion;
        public Microsoft.Office.Tools.CustomTaskPane suggestionTaskPane;
        private int maxPhraseLength = 12;
        private bool isReplacing = false;
        public bool isAbbreviationEnabled = true;
        private bool frozeSuggestions = false;

        private string lastReplacedShortForm = "";
        private string lastReplacedFullForm = "";
        private bool justUndone = false;
        private List<(string Word, string Replacement)> _phraseCache;
        private System.Windows.Forms.Timer debounceTimer;
        private const int DebounceDelayMs = 300;
        private string lastUndoneWord = null;
        public int lastDocumentCharCount = -1;
        public bool isTypingDetected = false;


        private Trie trie = new Trie();

        private List<string> allPhrases;


        private string lastWord = "";
        private Timer typingTimer;
        internal SuggestionPaneControl SuggestionPaneControl;
        bool replaceAllChosen = false;
        bool ignoreAllChosen = false;
        private bool replaceAllForPhrase;
        private bool ignoreAllForPhrase;
        public Dictionary<Word.Window, CustomTaskPane> taskPanes = new Dictionary<Word.Window, CustomTaskPane>();
        public HashSet<Word.Window> userClosedTaskPanes = new HashSet<Word.Window>();
        public HashSet<Word.Window> taskPaneOpenedOnce = new HashSet<Word.Window>();
        public bool suggestionFROMInput = false;

        private HashSet<string> _globallyReplacedPhrases = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> IgnoredAbbreviations = new HashSet<string>(StringComparer.InvariantCultureIgnoreCase);

        public bool isProgrammaticChange = false;
        public int lastSelectionStart = -1;



        internal class Win32WindowWrapper : IWin32Window
        {
            private readonly IntPtr _handle;

            public Win32WindowWrapper(IntPtr handle)
            {
                _handle = handle;
            }

            public IntPtr Handle => _handle;
        }



        private void TrackTaskPaneVisibility(CustomTaskPane pane, Word.Window window)
        {
            pane.VisibleChanged += (s, e) =>
            {
                if (!pane.Visible)
                {
                    userClosedTaskPanes.Add(window);
                    taskPaneOpenedOnce.Remove(window); // allow future reopen if needed
                    if (debug) { 
                        System.Diagnostics.Debug.WriteLine("[DEBUG] User closed TaskPane.");
                    }
                }
                else
                {
                    userClosedTaskPanes.Remove(window);
                    if (debug)
                    {
                        System.Diagnostics.Debug.WriteLine("[DEBUG] TaskPane shown.");
                    }
                }
            };
        }



        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                
                if (debug)
                {
                    System.Windows.Forms.MessageBox.Show(
                    "lastLoadedVersion" + lastLoadedVersion + "currentVersion" + currentVersion,
                    "Abbreviation Loading status",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information
                 );
                }

                var autoCorrect = Globals.ThisAddIn.Application.AutoCorrect;
                for (int i = autoCorrect.Entries.Count; i >= 1; i--)
                {
                    autoCorrect.Entries[i].Delete();
                }

              


                if (lastLoadedVersion != currentVersion)
                {
                    // Version changed → clear file cache
                    if (debug)
                    {
                        System.Windows.Forms.MessageBox.Show(
                        "Clear cache because lastLoadedVersion" + lastLoadedVersion + "currentVersion" + currentVersion,
                        "Abbreviation Loading status",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information
                     );
                    }
                    reloadAbbrDataFromDict = true;
                    AbbreviationManager.ClearCacheFile();
                    Properties.Settings.Default.IsAutoCorrectLoaded = false;
                    Properties.Settings.Default.LastLoadedAbbreviationVersion = currentVersion;
                    Properties.Settings.Default.Save();
                    Properties.Settings.Default.Reload();
                }


                AbbreviationManager.LoadAbbreviations();

                allPhrases = AbbreviationManager.GetAllPhrases().ToList();

                trie = new Trie();
                foreach (var phrase in allPhrases)
                {
                    trie.Insert(phrase.ToLowerInvariant());
                }

                loadAllAbbreviaitons();

                SuggestionPaneControl = new SuggestionPaneControl();
                suggestionTaskPane = this.CustomTaskPanes.Add(SuggestionPaneControl, "Abbreviation Suggestions");
                SuggestionPaneControl.OnTextChanged += SuggestionPaneControl_OnTextChanged;
                SuggestionPaneControl.OnSuggestionAccepted += SuggestionPaneControl_OnSuggestionAccepted;
                suggestionTaskPane.Width = 1000;
                //suggestionTaskPane.Visible = true;

                typingTimer = new Timer { Interval = 300 };
                typingTimer.Tick += TypingTimer_Tick;
                typingTimer.Start();

                debounceTimer = new System.Windows.Forms.Timer();
                debounceTimer.Interval = DebounceDelayMs;
                debounceTimer.Tick += DebounceTimer_Tick;

                ((Word.ApplicationEvents4_Event)this.Application).NewDocument += Application_NewDocument;
                //((Word.ApplicationEvents4_Event)this.Application).DocumentOpen += Application_DocumentOpen;
                ((Word.ApplicationEvents4_Event)this.Application).WindowActivate += Application_WindowActivate;
                ((Word.ApplicationEvents4_Event)this.Application).DocumentChange += Application_DocumentChange;

                if (this.Application.Documents.Count > 0 && this.Application.ActiveWindow != null)
                {
                    EnsureTaskPaneVisible(this.Application.ActiveWindow, "Startup");
                }


            }
            catch (Exception ex)
            {
                if (debug)
                {
                    System.Windows.Forms.MessageBox.Show(
                    "Error during startup: " + ex.Message,
                    "Startup Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning
                );
                }
            }
        }


        //private void Application_DocumentOpen(Word.Document Doc)
        //{
        //    if (!userClosedTaskPanes.Contains(this.Application.ActiveWindow))
        //        EnsureTaskPaneVisible(this.Application.ActiveWindow);
        //}

        public void RestartTypingTimer()
        {
            try
            {
                typingTimer?.Stop();
                typingTimer?.Dispose();

                typingTimer = new System.Windows.Forms.Timer
                {
                    Interval = 300
                };

                typingTimer.Tick += TypingTimer_Tick;
                typingTimer.Start();
            }
            catch
            {
                // Never crash Word
            }
        }


        private void Application_NewDocument(Word.Document Doc)
        {
            try
            {
                if (debug)
                {
                    MessageBox.Show("New document created. Name: " + Doc.Name, "Debug - NewDocument");
                }

                EnsureTaskPaneVisible(this.Application.ActiveWindow, "New Documnet");

                // Debug: Show how many phrases we currently have
                var phrases = AbbreviationManager.GetAllPhrases();
                if (phrases != null && phrases.Any())
                {
                    if (debug)
                    {
                        MessageBox.Show("Loaded phrases count: " + phrases.Count(), "Debug - Phrases Loaded");
                    }
                }
                else
                {
                    if (debug)
                    {
                        MessageBox.Show("No phrases found when creating new document.", "Debug - Phrases Missing");
                    }
                }

                loadAllAbbreviaitons();
            }
            catch (Exception ex)
            {
                if (debug)
                {
                    MessageBox.Show("Error in Application_NewDocument: " + ex.Message, "Error");
                }
            }
        }


        private void Application_WindowActivate(Word.Document Doc, Word.Window Wn)
        {
            EnsureTaskPaneVisible(Wn,"windowActivate");
        }

        private void Application_DocumentChange()
        {
            if (this.Application.Documents.Count == 0)
            {
                typingTimer?.Stop();
                debounceTimer?.Stop();
            }
            else
            {
                typingTimer?.Start();
                debounceTimer?.Start();
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
                    if (debug)
                    {
                        System.Windows.Forms.MessageBox.Show(
                               "Loading Latest Abbreviation",
                               "AutoCorrect Entries",
                               System.Windows.Forms.MessageBoxButtons.OK,
                               System.Windows.Forms.MessageBoxIcon.Information
                           );
                    }
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
                        if (debug)
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
            } 

            Properties.Settings.Default.IsAutoCorrectLoaded = true;
            Properties.Settings.Default.Save();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //AbbreviationManager.ClearAutoCorrectCache();

        }




        //public void ToggleAbbreviationReplacement(bool enable)
        //{
        //    if (!Properties.Settings.Default.IsAutoCorrectLoaded)
        //    {
        //        reloadAbbrDataFromDict = enable;
        //        if (reloadAbbrDataFromDict)
        //        {
        //            isAbbreviationEnabled = true;
        //            AbbreviationManager.InitializeAutoCorrectCache(this.Application.AutoCorrect);
        //            if (suggestionTaskPane != null)
        //            {
        //                suggestionTaskPane.Visible = true;
        //            }
        //            System.Windows.Forms.MessageBox.Show("Abbreviation Replacement Enabled", "Status");
        //        }
        //        else
        //        {
        //            isAbbreviationEnabled = false;
        //            AbbreviationManager.ClearAutoCorrectCache();
        //            if (suggestionTaskPane != null)
        //            {
        //                suggestionTaskPane.Visible = false;
        //            }
        //            System.Windows.Forms.MessageBox.Show("Abbreviation Replacement Disabled", "Status");
        //        }
        //    }

        //}

        public bool HasAbbreviationsInDocument()
        {
            try
            {
                Word.Application app = this.Application;
                Word.Document doc = app?.ActiveDocument;
                if (doc == null)
                    return false;

                // Read document text ONCE
                string fullText = doc.Content.Text;
                if (string.IsNullOrWhiteSpace(fullText))
                    return false;

                // Get all abbreviation phrases
                var phrases = AbbreviationManager.GetAllPhrases()
                    .Where(p => !string.IsNullOrWhiteSpace(p))
                    .ToList();

                if (phrases.Count == 0)
                    return false;

                // Build regex pattern (escaped)
                string pattern = string.Join("|", phrases.Select(Regex.Escape));

                // Fast existence check (no allocations like Matches)
                return Regex.IsMatch(
                    fullText,
                    pattern,
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant
                );
            }
            catch
            {
                // Never allow a pre-check to crash Word
                return false;
            }
        }


        public void ToggleAbbreviationReplacement(bool enable)
        {
            isAbbreviationEnabled = enable;

            var window = this.Application.ActiveWindow;

            if (!taskPanes.ContainsKey(window))
            {
                SuggestionPaneControl = new SuggestionPaneControl();
                var pane = this.CustomTaskPanes.Add(SuggestionPaneControl, "Abbreviation Suggestions", window);
                pane.Width = 700;

                taskPanes[window] = pane;

                TrackTaskPaneVisibility(pane, window);
            }

            var currentPane = taskPanes[window];

            if (enable)
            {
                AbbreviationManager.InitializeAutoCorrectCache(this.Application.AutoCorrect);
                currentPane.Visible = true;
                if (debug)
                {
                    MessageBox.Show("Abbreviation Replacement Enabled", "Status");
                }
            }
            else
            {
                AbbreviationManager.ClearAutoCorrectCache();
                currentPane.Visible = false;
                if (debug)
                {
                    MessageBox.Show("Abbreviation Replacement Disabled", "Status");
                }
            }
        }





        //public void ReplaceAllDirectAbbreviations_Fast()
        //{
        //    _globallyReplacedPhrases.Clear();
        //    Word.Document doc = null;

        //    try
        //    {
        //        doc = this.Application.ActiveDocument;
        //        if (doc == null) return;

        //        this.Application.ScreenUpdating = false;
        //        this.Application.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
        //        this.Application.ShowAnimation = false;
        //        this.Application.DisplayStatusBar = false;
        //        this.Application.Options.ReplaceSelection = false;

        //        if (reloadAbbrDataFromDict)
        //        {
        //            AbbreviationManager.InitializeAutoCorrectCache(this.Application.AutoCorrect);
        //        }

        //        // 🔹 Longest phrases first (prevents partial replacements)
        //        var phrases = AbbreviationManager.GetAllPhrases()
        //                                         .Where(p => !string.IsNullOrWhiteSpace(p))
        //                                         .OrderByDescending(p => p.Length)
        //                                         .ToList();

        //        foreach (var phrase in phrases)
        //        {
        //            string replacement =
        //                AbbreviationManager.GetFromAutoCorrectCache(phrase)
        //                ?? AbbreviationManager.GetAbbreviation(phrase);

        //            if (string.IsNullOrWhiteSpace(replacement))
        //                continue;

        //            Word.Find find = doc.Content.Find;
        //            find.ClearFormatting();
        //            find.Replacement.ClearFormatting();

        //            find.Text = phrase;
        //            find.Replacement.Text = replacement;

        //            find.MatchCase = false;
        //            find.MatchWholeWord = true;     // ✅ safe whole phrase
        //            find.MatchWildcards = false;
        //            find.Wrap = Word.WdFindWrap.wdFindContinue;

        //            // 🚀 Single ReplaceAll per phrase (Word-native, safe)
        //            bool replaced = find.Execute(
        //                Replace: Word.WdReplace.wdReplaceAll,
        //                Forward: true
        //            );

        //            if (replaced)
        //            {
        //                _globallyReplacedPhrases.Add(phrase);
        //            }

        //        }
        //    }
        //    finally
        //    {
        //        this.Application.ScreenUpdating = true;
        //        this.Application.DisplayAlerts = Word.WdAlertLevel.wdAlertsAll;
        //        this.Application.ShowAnimation = true;
        //        this.Application.DisplayStatusBar = true;
        //        this.Application.Options.ReplaceSelection = true;
        //    }
        //}

        public bool HasAnyAbbreviations()
        {
            Word.Document doc = this.Application.ActiveDocument;
            if (doc == null) return false;

            string text = doc.Content.Text;
            if (string.IsNullOrWhiteSpace(text)) return false;

            var phrases = AbbreviationManager.GetAllPhrases()
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .Distinct(StringComparer.InvariantCultureIgnoreCase)
                .OrderByDescending(p => p.Length)
                .Select(Regex.Escape);

            if (!phrases.Any()) return false;

            string pattern = string.Join("|", phrases);

            return Regex.IsMatch(
                text,
                pattern,
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant
            );
        }


        //public void ReplaceAllDirectAbbreviations_Fast()
        //{
        //    ReplaceAllAbbreviations();
        //    Word.Application app = this.Application;
        //    Word.Document doc = app.ActiveDocument;
        //    if (doc == null) return;

        //    Word.Range originalSelection = null;

        //    try
        //    {
        //        if (reloadAbbrDataFromDict)
        //            AbbreviationManager.InitializeAutoCorrectCache(app.AutoCorrect);

        //        Word.Range contentRange = doc.Content;
        //        string originalText = contentRange.Text;

        //        var map = AbbreviationManager.GetAllPhrases()
        //            .Where(p => !string.IsNullOrWhiteSpace(p))
        //            .Distinct(StringComparer.InvariantCultureIgnoreCase)
        //            .ToDictionary(
        //                p => p,
        //                p => AbbreviationManager.GetFromAutoCorrectCache(p)
        //                     ?? AbbreviationManager.GetAbbreviation(p),
        //                StringComparer.InvariantCultureIgnoreCase
        //            );

        //        if (map.Count == 0)
        //        {
        //            WordDialogHelper.ShowInfo("No abbreviations are configured.");
        //            return;
        //        }

        //        string pattern = string.Join("|",
        //            map.Keys
        //               .OrderByDescending(k => k.Length)
        //               .Select(Regex.Escape)
        //        );

        //        var regex = new Regex(
        //            pattern,
        //            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant
        //        );

        //        string replacedText = regex.Replace(originalText, m =>
        //        {
        //            string key = m.Value;
        //            return map.TryGetValue(key, out var repl)
        //                ? repl
        //                : key;
        //        });

        //        if (replacedText == originalText)
        //        {
        //            WordDialogHelper.ShowInfo(
        //                "No abbreviations were found in the document."
        //            );
        //            return;
        //        }

        //        app.ScreenUpdating = false;
        //        app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
        //        app.ShowAnimation = false;

        //        originalSelection = app.Selection.Range.Duplicate;
        //        app.UndoRecord.StartCustomRecord("Replace All Abbreviations");

        //        contentRange.Text = replacedText;
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(
        //            "An error occurred while replacing abbreviations.\n\n" + ex.Message,
        //            "Abbreviation Tool",
        //            MessageBoxButtons.OK,
        //            MessageBoxIcon.Error
        //        );
        //    }
        //    finally
        //    {
        //        try { app.UndoRecord.EndCustomRecord(); } catch { }

        //        originalSelection?.Select();
        //        app.ScreenUpdating = true;
        //        app.DisplayAlerts = Word.WdAlertLevel.wdAlertsAll;
        //        app.ShowAnimation = true;
        //    }
        //}

        //    private void ReplaceInRange(
        //Word.Range sourceRange,
        //Regex regex,
        //Dictionary<string, string> map)
        //    {
        //        foreach (Word.Paragraph para in sourceRange.Paragraphs)
        //        {
        //            Word.Range paraRange = para.Range;

        //            string originalText;
        //            try
        //            {
        //                originalText = paraRange.Text;
        //            }
        //            catch
        //            {
        //                continue; // Word returned invalid range
        //            }

        //            // 🔒 HARD NULL GUARD (VERY IMPORTANT)
        //            if (string.IsNullOrEmpty(originalText))
        //                continue;

        //            // Remove paragraph mark
        //            if (originalText.EndsWith("\r"))
        //            {
        //                paraRange.End -= 1;
        //                originalText = originalText.Substring(0, originalText.Length - 1);
        //            }

        //            // Remove table cell marker
        //            if (originalText.EndsWith("\a"))
        //            {
        //                paraRange.End -= 1;
        //                originalText = originalText.Substring(0, originalText.Length - 1);
        //            }

        //            if (string.IsNullOrWhiteSpace(originalText))
        //                continue;

        //            string replacedText = regex.Replace(
        //                originalText,
        //                match =>
        //                {
        //                    if (map.TryGetValue(match.Value, out string repl))
        //                        return repl;

        //                    return match.Value;
        //                });

        //            // Write back only if changed
        //            if (!string.Equals(originalText, replacedText, StringComparison.Ordinal))
        //            {
        //                paraRange.Text = replacedText;
        //            }
        //        }
        //    }

        private void ReplaceInRange(
            Word.Range sourceRange,
            Regex regex,
            Dictionary<string, string> map)
        {
            string text;
            try
            {
                text = sourceRange.Text;
            }
            catch
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(text))
                return;

            MatchCollection matches = regex.Matches(text);
            if (matches.Count == 0)
                return;

            for (int i = matches.Count - 1; i >= 0; i--)
            {
                Match m = matches[i];

                if (!map.TryGetValue(m.Value, out string repl))
                    continue;

                int start = sourceRange.Start + m.Index;
                int end = start + m.Length;

                if (end > sourceRange.End)
                    continue;

                Word.Range r = sourceRange.Document.Range(start, end);

                if (r.Text.Contains("\a") || r.Text.Contains("\r"))
                    continue;

                r.Text = repl;
            }
        }





        public void ReplaceAllDirectAbbreviations_Fast()
        {
            Word.Application app = this.Application;
            Word.Document doc = app.ActiveDocument;
            if (doc == null) return;

            if (reloadAbbrDataFromDict)
                AbbreviationManager.InitializeAutoCorrectCache(app.AutoCorrect);

            var ignored = this.IgnoredAbbreviations;

            var map = AbbreviationManager.GetAllPhrases()
                .Where(p =>
                    !string.IsNullOrWhiteSpace(p) &&
                    !ignored.Contains(p)
                )
                .Distinct(StringComparer.InvariantCultureIgnoreCase)
                .ToDictionary(
                    p => p,
                    p => AbbreviationManager.GetFromAutoCorrectCache(p)
                          ?? AbbreviationManager.GetAbbreviation(p),
                    StringComparer.InvariantCultureIgnoreCase
                );


            if (map.Count == 0)
            {
                WordDialogHelper.ShowInfo("No abbreviations are configured.");
                return;
            }

            string pattern = @"\b(" +
                string.Join("|",
                    map.Keys
                       .OrderByDescending(k => k.Length)
                       .Select(Regex.Escape)
                ) +
                @")\b";


            Regex regex = new Regex(
                pattern,
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant
            );

            app.ScreenUpdating = false;
            app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            app.ShowAnimation = false;
            app.UndoRecord.StartCustomRecord("Replace All Abbreviations");

            try
            {
                foreach (Word.Paragraph para in doc.Paragraphs)
                {
                    if (!para.Range.Information[Word.WdInformation.wdWithInTable])
                        ReplaceInRange(para.Range, regex, map);
                }

                foreach (Word.Table table in doc.Tables)
                {
                    foreach (Word.Row row in table.Rows)
                    {
                        foreach (Word.Cell cell in row.Cells)
                        {
                            ReplaceInRange(cell.Range, regex, map);
                        }
                    }
                }
            }
            finally
            {
                try { app.UndoRecord.EndCustomRecord(); } catch { }

                app.ScreenUpdating = true;
                app.DisplayAlerts = Word.WdAlertLevel.wdAlertsAll;
                app.ShowAnimation = true;
                Globals.ThisAddIn.IgnoredAbbreviations.Clear();

            }
        }


















        public void ReplaceAllDirectAbbreviations()
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

                        // 🚫 Disable UI flicker
                        this.Application.ScreenUpdating = false;
                        this.Application.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                        this.Application.ShowAnimation = false;
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

                    var phrases = AbbreviationManager.GetAllPhrases();
                    int total = phrases.Count;
                    int processed = 0;

                    foreach (var phrase in phrases)
                    {
                        processed++;
                        int percent = (processed * 100) / total;
                        progressForm.UpdateProgress(percent, $"Replacing '{phrase}' ({processed}/{total})...");

                        string replacement = AbbreviationManager.GetFromAutoCorrectCache(phrase)
                                           ?? AbbreviationManager.GetAbbreviation(phrase);

                        syncContext.Send(_ =>
                        {
                            try
                            {
                                Word.Find find = doc.Content.Find;
                                find.ClearFormatting();
                                find.Text = phrase;
                                find.MatchCase = false;
                                find.MatchWholeWord = true;
                                find.MatchWildcards = false;
                                find.Replacement.ClearFormatting();
                                find.Replacement.Text = replacement;

                                // 🚀 One-shot ReplaceAll
                                find.Execute(
                                    Replace: Word.WdReplace.wdReplaceAll,
                                    Forward: true,
                                    Wrap: Word.WdFindWrap.wdFindContinue
                                );

                                // Log to Form1 (caret hidden)
                            }
                            catch (Exception ex)
                            {
                                processError = ex;
                                completed = true;
                            }
                        }, null);

                        if (completed) break;
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
                        // ✅ Restore UI
                        this.Application.ScreenUpdating = true;
                        this.Application.DisplayAlerts = Word.WdAlertLevel.wdAlertsAll;
                        this.Application.ShowAnimation = true;
                        this.Application.DisplayStatusBar = true;
                        this.Application.Options.ReplaceSelection = true;
                    }, null);

                    completed = true;
                    syncContext.Post(_ => progressForm.Close(), null);
                }
            });

            progressThread.Start();
            progressForm.ShowDialog();

            if (processError != null)
            {
                throw processError;
            }
        }

        public void ReplaceAllReverseAbbreviations()
        {
            if (!isAbbreviationEnabled) return;
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

                    // ✅ Build reverse dictionary (Replacement -> Phrase)
                    var phrases = AbbreviationManager.GetAllPhrases();
                    var reverseMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                    foreach (var phrase in phrases)
                    {
                        string replacement = AbbreviationManager.GetFromAutoCorrectCache(phrase)
                                             ?? AbbreviationManager.GetAbbreviation(phrase);

                        if (!string.IsNullOrWhiteSpace(replacement) &&
                            !reverseMap.ContainsKey(replacement))
                        {
                            reverseMap[replacement] = phrase; // reverse mapping
                        }
                    }

                    int total = reverseMap.Count;
                    int processed = 0;

                    foreach (var kvp in reverseMap)
                    {
                        processed++;
                        int percent = (processed * 100) / total;
                        progressForm.UpdateProgress(percent, $"Reversing '{kvp.Key}' → '{kvp.Value}' ({processed}/{total})...");

                        syncContext.Send(_ =>
                        {
                            try
                            {
                                Word.Find find = doc.Content.Find;
                                find.ClearFormatting();
                                find.Text = kvp.Key;  // search replacement
                                find.MatchCase = false;
                                find.MatchWholeWord = true;
                                find.MatchWildcards = false;

                                find.Replacement.ClearFormatting();
                                find.Replacement.Text = kvp.Value; // replace with original phrase

                                // 🚀 ReplaceAll in one go
                                find.Execute(
                                    Replace: Word.WdReplace.wdReplaceAll,
                                    Forward: true,
                                    Wrap: Word.WdFindWrap.wdFindContinue
                                );
                            }
                            catch (Exception ex)
                            {
                                processError = ex;
                                completed = true;
                            }
                        }, null);

                        if (completed) break;
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

            if (processError != null)
            {
                throw processError;
            }
        }



        //public void ReplaceAllAbbreviations()
        //{
        //    if (!isAbbreviationEnabled) return;
        //    var progressForm = new ProgressForm();

        //    var syncContext = System.Threading.SynchronizationContext.Current;
        //    bool completed = false;
        //    Exception processError = null;

        //    var progressThread = new System.Threading.Thread(() =>
        //    {
        //        try
        //        {
        //            Word.Document doc = null;
        //            syncContext.Send(_ =>
        //            {
        //                doc = this.Application.ActiveDocument;
        //                this.Application.ScreenUpdating = false; 
        //                this.Application.DisplayStatusBar = false; 
        //                this.Application.Options.ReplaceSelection = false; 
        //            }, null);

        //            if (reloadAbbrDataFromDict)
        //            {
        //                syncContext.Send(_ =>
        //                {
        //                    AbbreviationManager.InitializeAutoCorrectCache(this.Application.AutoCorrect);
        //                }, null);
        //            }

        //            int totalWords = 0;
        //            syncContext.Send(_ =>
        //            { 
        //                totalWords = doc.Words.Count;
        //            }, null);

        //            int totalChunks = (totalWords + CHUNK_SIZE - 1) / CHUNK_SIZE;
        //            int currentChunk = 0;


        //            for (int startIndex = 1; startIndex <= totalWords && !completed; startIndex += CHUNK_SIZE)
        //            {
        //                currentChunk++;
        //                int endIndex = Math.Min(startIndex + CHUNK_SIZE - 1, totalWords);

        //                int percentage = (currentChunk * 100) / totalChunks;
        //                progressForm.UpdateProgress(percentage, $"Processing chunk {currentChunk} of {totalChunks}...");

        //                syncContext.Send(_ =>
        //                {
        //                    try
        //                    {
        //                        Word.Range chunkRange = doc.Range(doc.Words[startIndex].Start, doc.Words[endIndex].End);
        //                        string chunkText = chunkRange.Text;
        //                        bool hasMatches = false;

        //                        foreach (var phrase in AbbreviationManager.GetAllPhrases())
        //                        {
        //                            if (chunkText.IndexOf(phrase, StringComparison.OrdinalIgnoreCase) != -1)
        //                            {
        //                                hasMatches = true;
        //                                break;
        //                            }
        //                        }

        //                        if (hasMatches)
        //                        {
        //                            foreach (var phrase in AbbreviationManager.GetAllPhrases())
        //                            {
        //                                string replacement = AbbreviationManager.GetFromAutoCorrectCache(phrase)
        //                                    ?? AbbreviationManager.GetAbbreviation(phrase);


        //                                Word.Find find = doc.Content.Find;
        //                                find.ClearFormatting();
        //                                find.Text = phrase;
        //                                find.MatchCase = false;
        //                                find.MatchWholeWord = true;
        //                                find.Wrap = Word.WdFindWrap.wdFindStop;



        //                                while (find.Execute())
        //                                {
        //                                    Word.Range matchRange = find.Parent as Word.Range;
        //                                    if (matchRange == null) break;


        //                                    if (ignoreAllForPhrase)
        //                                    {
        //                                        continue;
        //                                    }

        //                                    using (var dlg = new ReplaceDialog(phrase, replacement))
        //                                    {
        //                                        var result = dlg.ShowDialog();
        //                                        if (result == DialogResult.OK)
        //                                        {
        //                                            switch (dlg.UserChoice)
        //                                            {
        //                                                case ReplaceDialog.ReplaceAction.Replace:
        //                                                    matchRange.Text = replacement;
        //                                                    break;

        //                                                case ReplaceDialog.ReplaceAction.ReplaceAll:
        //                                                    matchRange.Text = replacement;
        //                                                    ReplaceAllDirectAbbreviations_Fast();
        //                                                    return;

        //                                                case ReplaceDialog.ReplaceAction.Ignore:
        //                                                    // skip this one
        //                                                    break;

        //                                                case ReplaceDialog.ReplaceAction.IgnoreAll:
        //                                                    ignoreAllForPhrase = true; // future occurrences skipped automatically
        //                                                    break;

        //                                                case ReplaceDialog.ReplaceAction.Cancel:
        //                                                case ReplaceDialog.ReplaceAction.Close:
        //                                                    return; // stop everything
        //                                            }
        //                                        }
        //                                        else
        //                                        {
        //                                            return; // user closed dialog abruptly → stop everything
        //                                        }
        //                                    }
        //                                }



        //                            }
        //                        }

        //                        if (chunkRange != null)
        //                            System.Runtime.InteropServices.Marshal.ReleaseComObject(chunkRange);
        //                    }
        //                    catch (Exception ex)
        //                    {
        //                        processError = ex;
        //                        completed = true; 
        //                    }
        //                }, null);
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            processError = ex;
        //        }
        //        finally
        //        {
        //            syncContext.Send(_ =>
        //            {
        //                this.Application.ScreenUpdating = true; 
        //                this.Application.DisplayStatusBar = true; 
        //                this.Application.Options.ReplaceSelection = true; 
        //            }, null);

        //            completed = true;
        //            syncContext.Post(_ => progressForm.Close(), null);
        //        }
        //    });

        //    progressThread.Start();
        //    progressForm.ShowDialog();

        //}


       

        public class MatchResult
        {
            public string Phrase { get; set; }
            public string Replacement { get; set; }
            public int StartIndex { get; set; }
            public int Length { get; set; }
        }

        public List<MatchResult> CollectAllAbbreviations()
        {
            var results = new List<MatchResult>();
            Word.Document doc = this.Application.ActiveDocument;
            string fullText = doc.Content.Text;

            // Get all phrases, longest first
            var phrases = AbbreviationManager.GetAllPhrases()
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .OrderByDescending(p => p.Length)
                .ToList();

            var usedIndexes = new HashSet<int>();

            foreach (var phrase in phrases)
            {
                if (_globallyReplacedPhrases.Contains(phrase))
                    continue;

                string pattern = $@"\b{Regex.Escape(phrase)}\b";

                var matches = Regex.Matches(
                    fullText,
                    pattern,
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant
                );

                foreach (Match m in matches)
                {
                    // Skip overlapping ranges (longest phrase wins)
                    bool overlap = false;
                    for (int i = m.Index; i < m.Index + m.Length; i++)
                    {
                        if (usedIndexes.Contains(i))
                        {
                            overlap = true;
                            break;
                        }
                    }
                    if (overlap)
                        continue;

                    string replacement =
                        AbbreviationManager.GetFromAutoCorrectCache(phrase)
                        ?? AbbreviationManager.GetAbbreviation(phrase);

                    results.Add(new MatchResult
                    {
                        Phrase = phrase,
                        Replacement = replacement,
                        StartIndex = m.Index,
                        Length = m.Length
                    });

                    // Mark indexes as used
                    for (int i = m.Index; i < m.Index + m.Length; i++)
                        usedIndexes.Add(i);
                }
            }

            // ✅ RETURN UNIQUE PHRASES ONLY (one row per phrase)
            return results
                .GroupBy(r => r.Phrase, StringComparer.OrdinalIgnoreCase)
                .Select(g =>
                {
                    var first = g.OrderBy(x => x.StartIndex).First();
                    return new MatchResult
                    {
                        Phrase = first.Phrase,
                        Replacement = first.Replacement,
                        StartIndex = first.StartIndex,
                        Length = first.Length
                    };
                })
                .OrderBy(r => r.StartIndex)
                .ToList();
        }




        public SuggestionPaneControl EnsureTaskPaneVisible(Word.Window window, string from)
        {
            if (window == null || this.Application.Documents.Count == 0)
                return null;

            // Skip if user manually closed
            if (userClosedTaskPanes.Contains(window))
                return null;

            // Check if pane already exists
            if (taskPanes.TryGetValue(window, out var existingPane))
                return existingPane.Control as SuggestionPaneControl;

            // Check CustomTaskPanes collection
            foreach (CustomTaskPane pane in this.CustomTaskPanes)
            {
                if (pane.Window == window && pane.Title == "Abbreviation Suggestions")
                {
                    taskPanes[window] = pane;
                    return pane.Control as SuggestionPaneControl;
                }
            }

            // Create a new pane
            var control = new SuggestionPaneControl();
            control.OnTextChanged += SuggestionPaneControl_OnTextChanged;
            control.OnSuggestionAccepted += SuggestionPaneControl_OnSuggestionAccepted;

            var newPane = this.CustomTaskPanes.Add(control, "Abbreviation Suggestions", window);
            newPane.Width = 1000;
            taskPanes[window] = newPane;

            TrackTaskPaneVisibility(newPane, window);

            // Do NOT set Visible = true here — leave it to the caller
            return control;
        }







        public void ReplaceAllAbbreviations()
        {
            try
            {
                Word.Window activeWindow = this.Application.ActiveWindow;

                // Ensure task pane is visible, even if closed
                var paneControl = EnsureTaskPaneVisible(activeWindow, "Replace ALL");

                // 🔹 If the pane is not found, create or reopen it
                if (paneControl == null)
                {
                    if (!taskPanes.ContainsKey(activeWindow))
                    {
                        SuggestionPaneControl = new SuggestionPaneControl();
                        var pane = this.CustomTaskPanes.Add(SuggestionPaneControl, "Abbreviation Suggestions", activeWindow);
                        pane.Width = 1000;
                        taskPanes[activeWindow] = pane;
                        TrackTaskPaneVisibility(pane, activeWindow);
                    }

                    // Retrieve the control and show pane
                    var currentPane = taskPanes[activeWindow];
                    SuggestionPaneControl = currentPane.Control as SuggestionPaneControl;
                    currentPane.Visible = true;
                    paneControl = SuggestionPaneControl;
                }
                else
                {
                    // 🔹 If pane exists but is hidden, show it
                    var pane = taskPanes.ContainsKey(activeWindow) ? taskPanes[activeWindow] : null;
                    if (pane != null && !pane.Visible)
                        pane.Visible = true;
                }

                var matches = CollectAllAbbreviations();

                if (matches == null || matches.Count == 0)
                {
                    WordDialogHelper.ShowInfo(
                         "No words found in the document for which abbreviations are applicable."
                     );


                    paneControl?.Invoke(new Action(() =>
                    {
                        paneControl.LoadMatches(new List<MatchResult>());
                    }));

                    return;
                }


                if (paneControl.InvokeRequired)
                    paneControl.Invoke(new Action(() => paneControl.LoadMatches(matches)));
                else
                    paneControl.LoadMatches(matches);
            }
            catch (Exception ex)
            {
                if (debug)
                {
                    MessageBox.Show(
                        "Error in ReplaceAllAbbreviations: " + ex.Message + "\n" + ex.StackTrace,
                        "Exception",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error
                    );
                }
            }
        }



        private void ReplaceFirstInRange(Word.Document doc, string search, string replace)
        {
            Word.Range rng = doc.Content;
            var find = rng.Find;
            find.ClearFormatting();
            find.Text = search;
            find.Replacement.ClearFormatting();
            find.Replacement.Text = replace;

            find.Forward = true;
            find.Wrap = Word.WdFindWrap.wdFindStop;
            find.MatchWholeWord = true;   // <-- exact match only

            // Replace once and show immediately
            doc.Application.ScreenUpdating = true;
            find.Execute(Replace: Word.WdReplace.wdReplaceOne);
            doc.Application.ScreenUpdating = false;
        }

        public void MarkPhraseAsReplaced(string phrase)
        {
            if (!string.IsNullOrWhiteSpace(phrase))
                _globallyReplacedPhrases.Add(phrase);
        }


        public void ReplaceAbbreviation(string word, string replacement, bool replaceAll = false)
        {
            var app = this.Application;
            var doc = app.ActiveDocument;

            if (doc == null)
            {
                if (debug)
                {
                    MessageBox.Show("No active document found.");
                }
                return;
            }

            try
            {
                Word.Range range = doc.Content;
                Word.Find find = range.Find;

                find.ClearFormatting();
                find.Replacement.ClearFormatting();

                find.Text = word;
                find.Replacement.Text = replacement;

                find.MatchCase = false;
                find.MatchWholeWord = true;
                find.MatchWildcards = false;
                find.Forward = true;
                find.Wrap = WdFindWrap.wdFindContinue;

                // ✅ KEY FIX: Replace ONE or ALL correctly
                find.Execute(
                    Replace: replaceAll
                        ? WdReplace.wdReplaceAll
                        : WdReplace.wdReplaceOne
                );
            }
            catch (Exception ex)
            {
                if (debug)
                {
                    MessageBox.Show("Error replacing abbreviation: " + ex.Message);
                    System.Diagnostics.Debug.WriteLine("Error replacing abbreviation: " + ex.Message);
                }
            }
        }




        private void ReplaceAllInRange(Word.Document doc, string search, string replace)
        {
            Word.Range rng = doc.Content;
            rng.Find.Execute(FindText: search,
                             ReplaceWith: replace,
                             Replace: Word.WdReplace.wdReplaceAll,
                             MatchCase: false,
                             MatchWholeWord: true);
        }

        private void SkipFirstInRange(Word.Document doc, string search)
        {
            Word.Range rng = doc.Content;
            if (rng.Find.Execute(search, MatchCase: false, MatchWholeWord: true))
            {
                rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            }
        }

        private void RemoveAllInRange(Word.Document doc, string search)
        {
            Word.Range rng = doc.Content;
            rng.Find.Execute(FindText: search,
                             ReplaceWith: "",
                             Replace: Word.WdReplace.wdReplaceAll,
                             MatchCase: false,
                             MatchWholeWord: true);
        }


        //public void HighlightAllAbbreviations()
        //{
        //    if (!isAbbreviationEnabled) return;

        //    var progressForm = new ProgressForm();
        //    var syncContext = System.Threading.SynchronizationContext.Current;
        //    Exception processError = null;

        //    var highlightThread = new System.Threading.Thread(() =>
        //    {
        //        try
        //        {
        //            Word.Document doc = null;
        //            syncContext.Send(_ =>
        //            {
        //                doc = this.Application.ActiveDocument;
        //                this.Application.ScreenUpdating = false;
        //                this.Application.DisplayStatusBar = false;
        //                this.Application.Options.ReplaceSelection = false;
        //            }, null);

        //            if (!AbbreviationManager.IsAutoCorrectCacheInitialized())
        //            {
        //                syncContext.Send(_ =>
        //                {
        //                    AbbreviationManager.InitializeAutoCorrectCache(this.Application.AutoCorrect);
        //                }, null);
        //            }

        //            var phrases = AbbreviationManager.GetAllPhrases()
        //                                             .OrderByDescending(p => p.Length)
        //                                             .ToList();

        //            // Build one combined regex
        //            string pattern = string.Join("|", phrases.Select(p => $@"(?<!\w){Regex.Escape(p)}(?!\w)"));
        //            Regex regex = new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        //            string docText = null;
        //            syncContext.Send(_ => docText = doc.Content.Text, null);

        //            var matches = regex.Matches(docText);
        //            int totalMatches = matches.Count;
        //            int processed = 0;

        //            foreach (Match match in matches)
        //            {
        //                processed++;
        //                int percentage = (processed * 100) / totalMatches;
        //                progressForm.UpdateProgress(percentage, $"Highlighting {processed} of {totalMatches}...");

        //                int start = match.Index + 1; // Word ranges are 1-based
        //                int end = start + match.Length - 1;

        //                syncContext.Send(_ =>
        //                {
        //                    Word.Range range = doc.Range(start, end);
        //                    range.Font.Color = Word.WdColor.wdColorRed;
        //                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
        //                }, null);
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            processError = ex;
        //        }
        //        finally
        //        {
        //            syncContext.Send(_ =>
        //            {
        //                this.Application.ScreenUpdating = true;
        //                this.Application.DisplayStatusBar = true;
        //                this.Application.Options.ReplaceSelection = true;
        //                this.Application.Visible = true;
        //            }, null);

        //            syncContext.Post(_ => progressForm.Close(), null);
        //        }
        //    });

        //    highlightThread.Start();
        //    progressForm.ShowDialog();

        //    if (processError != null)
        //    {
        //        System.Windows.Forms.MessageBox.Show(
        //            "Error during highlighting: " + processError.Message,
        //            "Error",
        //            System.Windows.Forms.MessageBoxButtons.OK,
        //            System.Windows.Forms.MessageBoxIcon.Error
        //        );
        //    }
        //}

        public void HighlightAllAbbreviations()
        {

            if (!isAbbreviationEnabled) return;


            Word.Document doc = this.Application.ActiveDocument;

            if (!AbbreviationManager.IsAutoCorrectCacheInitialized())
            {
                AbbreviationManager.InitializeAutoCorrectCache(this.Application.AutoCorrect);
            }

            var phrases = AbbreviationManager.GetAllPhrases()
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .Where(p => !_globallyReplacedPhrases.Contains(p))
                .OrderByDescending(p => p.Length)
                .ToList();

            if (phrases.Count == 0)
            {
                WordDialogHelper.ShowInfo(
                        "No words found in the document for which abbreviations are applicable."
                    );
                return;
            }

            string pattern = @"(?<!\w)(" + string.Join("|", phrases.Select(Regex.Escape)) + @")(?!\w)";
            Regex regex = new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

            int totalMatches = 0;

            foreach (Word.Paragraph para in doc.Paragraphs)
            {
                string text = para.Range.Text;
                if (!string.IsNullOrWhiteSpace(text))
                    totalMatches += regex.Matches(text).Count;

                System.Runtime.InteropServices.Marshal.ReleaseComObject(para);
            }

            if (totalMatches == 0)
            {
                WordDialogHelper.ShowInfo(
                        "No words found in the document for which abbreviations are applicable."
                    );
                return;
            }


            var progressForm = new ProgressForm();
            var syncContext = System.Threading.SynchronizationContext.Current;

            Exception processError = null;

            var highlightThread = new System.Threading.Thread(() =>
            {
                try
                {
                    syncContext.Send(_ =>
                    {
                        this.Application.ScreenUpdating = false;
                        this.Application.DisplayStatusBar = false;
                        this.Application.Options.ReplaceSelection = false;
                    }, null);

                    int processed = 0;

                    foreach (Word.Paragraph para in doc.Paragraphs)
                    {
                        string paraText = para.Range.Text;
                        if (string.IsNullOrWhiteSpace(paraText)) continue;

                        var matches = regex.Matches(paraText);

                        foreach (Match match in matches)
                        {
                            if (_globallyReplacedPhrases.Contains(match.Value))
                                continue;

                            processed++;
                            int percentage = (processed * 100) / totalMatches;

                            syncContext.Send(_ =>
                            {
                                int start = para.Range.Start + match.Index;
                                int end = start + match.Length;

                                if (end > para.Range.End)
                                    end = para.Range.End;

                                Word.Range range = doc.Range(start, end);
                                range.Font.Color = Word.WdColor.wdColorRed;
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);

                                progressForm.UpdateProgress(
                                    percentage,
                                    $"Highlighting {processed} of {totalMatches}..."
                                );
                            }, null);
                        }

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(para);
                    }
                }
                finally
                {
                    syncContext.Send(_ =>
                    {
                        this.Application.ScreenUpdating = true;
                        this.Application.DisplayStatusBar = true;
                        this.Application.Options.ReplaceSelection = true;
                    }, null);

                    syncContext.Post(_ => progressForm.Close(), null);
                }
            });

            highlightThread.Start();
            progressForm.ShowDialog();
        }



        public void HighlightLike()
        {
            if (!isAbbreviationEnabled) return;

            var progressForm = new ProgressForm();
            var syncContext = System.Threading.SynchronizationContext.Current;
            Exception processError = null;
            var debugLog = new System.Text.StringBuilder();

            var highlightThread = new System.Threading.Thread(() =>
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

                    if (!AbbreviationManager.IsAutoCorrectCacheInitialized())
                    {
                        syncContext.Send(_ =>
                        {
                            AbbreviationManager.InitializeAutoCorrectCache(this.Application.AutoCorrect);
                        }, null);
                    }

                    // ✅ Stop words to ignore
                    var stopWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { "a", "an", "the", "of", "for", "to", "in", "on", "at", "by", "and", "or" };

                    // ✅ Build pattern parts for partial phrase matching
                    var phrases = AbbreviationManager.GetAllPhrases()
                        .Select(p =>
                        {
                            var words = p.Split(' ')
                                         .Where(w => w.Length > 2 && !stopWords.Contains(w))
                                         .ToList();

                            // ✅ Build progressive partial phrases: 
                            // e.g. "Accounting Manager Assistant" → "Accounting", "Accounting Manager", "Accounting Manager Assistant"
                            var parts = new List<string>();
                            for (int i = 1; i <= words.Count; i++)
                            {
                                var partial = string.Join(@"\s+", words.Take(i).Select(Regex.Escape));
                                parts.Add(partial);
                            }
                            return parts;
                        })
                        .SelectMany(p => p)
                        .Distinct()
                        .OrderByDescending(p => p.Length)
                        .ToList();

                    // ✅ Build final regex (now supports partials)
                    string pattern = string.Join("|", phrases);
                    Regex regex = new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

                    // ✅ Count matches
                    int totalMatches = 0;
                    syncContext.Send(_ =>
                    {
                        foreach (Word.Paragraph para in doc.Paragraphs)
                        {
                            string paraText = para.Range.Text;
                            if (!string.IsNullOrWhiteSpace(paraText))
                                totalMatches += regex.Matches(paraText)
                                    .Cast<Match>()
                                    .Count(m => !_globallyReplacedPhrases.Contains(m.Value));

                        }
                    }, null);

                    debugLog.AppendLine($"Total matches found: {totalMatches}");
                    int processed = 0;

                    // ✅ Highlight matches
                    foreach (Word.Paragraph para in doc.Paragraphs)
                    {
                        string paraText = para.Range.Text;
                        if (string.IsNullOrWhiteSpace(paraText)) continue;

                        var matches = regex.Matches(paraText);
                        foreach (Match match in matches)
                        {
                            processed++;
                            int percentage = (processed * 100) / (totalMatches == 0 ? 1 : totalMatches);

                            syncContext.Send(_ =>
                            {
                                try
                                {
                                    int start = para.Range.Start + match.Index;
                                    int end = start + match.Length;

                                    if (end > para.Range.End)
                                        end = para.Range.End;

                                    Word.Range range = doc.Range(start, end);
                                    range.Font.Color = Word.WdColor.wdColorRed;

                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                                }
                                catch (Exception ex)
                                {
                                    debugLog.AppendLine($"Error highlighting '{match.Value}': {ex.Message}");
                                }

                                progressForm.UpdateProgress(percentage, $"Highlighting {processed} of {totalMatches}...");
                            }, null);
                        }

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(para);
                    }
                }
                catch (Exception ex)
                {
                    processError = ex;
                    debugLog.AppendLine("❌ Exception: " + ex.ToString());
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

                    syncContext.Post(_ => progressForm.Close(), null);
                }
            });

            highlightThread.Start();
            progressForm.ShowDialog();
        }





        //public void HighlightAllAbbreviations()
        //{
        //    if (!isAbbreviationEnabled) return;
        //    var progressForm = new ProgressForm();
        //    var syncContext = System.Threading.SynchronizationContext.Current;
        //    bool completed = false;
        //    Exception processError = null;

        //    var progressThread = new System.Threading.Thread(() =>
        //    {
        //        try
        //        {
        //            Word.Document doc = null;
        //            syncContext.Send(_ =>
        //            {
        //                doc = this.Application.ActiveDocument;
        //                this.Application.ScreenUpdating = false;
        //                this.Application.DisplayStatusBar = false;
        //                this.Application.Options.ReplaceSelection = false;
        //            }, null);

        //            if (!AbbreviationManager.IsAutoCorrectCacheInitialized())
        //            {
        //                AbbreviationManager.InitializeAutoCorrectCache(this.Application.AutoCorrect);
        //            }

        //            int totalWords = 0;
        //            syncContext.Send(_ =>
        //            {
        //                totalWords = doc.Words.Count;
        //            }, null);

        //            int totalChunks = (totalWords + CHUNK_SIZE - 1) / CHUNK_SIZE;
        //            int currentChunk = 0;

        //            var phrases = AbbreviationManager.GetAllPhrases().ToList(); // cache once

        //            // Process document in chunks
        //            for (int startIndex = 1; startIndex <= totalWords && !completed; startIndex += CHUNK_SIZE)
        //            {
        //                currentChunk++;
        //                int endIndex = Math.Min(startIndex + CHUNK_SIZE - 1, totalWords);

        //                int percentage = (currentChunk * 100) / totalChunks;
        //                progressForm.UpdateProgress(percentage, $"Processing chunk {currentChunk} of {totalChunks}...");

        //                try
        //                {
        //                    Word.Range chunkRange = doc.Range(doc.Words[startIndex].Start, doc.Words[endIndex].End);
        //                    string chunkText = chunkRange.Text;

        //                    // only keep phrases that exist in this chunk
        //                    var matchingPhrases = phrases
        //                        .Where(p => chunkText.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0)
        //                        .ToList();

        //                    if (matchingPhrases.Count > 0)
        //                    {
        //                        foreach (var phrase in matchingPhrases)
        //                        {
        //                            //Word.Find find = chunkRange.Find;
        //                            //find.ClearFormatting();
        //                            //find.Text = phrase;
        //                            //find.MatchCase = false;
        //                            //find.MatchWholeWord = true;
        //                            //find.Wrap = Word.WdFindWrap.wdFindStop;

        //                            //find.Replacement.ClearFormatting();
        //                            //find.Replacement.Font.Color = Word.WdColor.wdColorRed;
        //                            //find.Replacement.Text = phrase;

        //                            //// do replacement inside chunk only (no wdFindContinue → prevents scanning full doc repeatedly)
        //                            //find.Execute(
        //                            //    ReplaceWith: phrase,
        //                            //    Replace: Word.WdReplace.wdReplaceAll,
        //                            //    MatchCase: false,
        //                            //    MatchWholeWord: true
        //                            //);

        //                            Word.Find find = chunkRange.Find;
        //                            find.ClearFormatting();
        //                            find.Text = phrase;
        //                            find.MatchCase = false;
        //                            find.MatchWholeWord = true;
        //                            find.Wrap = Word.WdFindWrap.wdFindStop;

        //                            while (find.Execute())
        //                            {
        //                                chunkRange.HighlightColorIndex = Word.WdColorIndex.wdYellow;
        //                                chunkRange.Font.Color = Word.WdColor.wdColorRed;

        //                                chunkRange.Start = chunkRange.End;
        //                                chunkRange.End = doc.Content.End;
        //                            }

        //                        }
        //                    }

        //                    // Release COM objects
        //                    if (chunkRange != null)
        //                        System.Runtime.InteropServices.Marshal.ReleaseComObject(chunkRange);
        //                }
        //                catch (Exception ex)
        //                {
        //                    processError = ex;
        //                    completed = true;
        //                }
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            processError = ex;
        //        }
        //        finally
        //        {
        //            syncContext.Send(_ =>
        //            {
        //                this.Application.ScreenUpdating = true;
        //                this.Application.DisplayStatusBar = true;
        //                this.Application.Options.ReplaceSelection = true;
        //            }, null);

        //            completed = true;
        //            syncContext.Post(_ => progressForm.Close(), null);
        //        }
        //    });

        //    progressThread.Start();
        //    progressForm.ShowDialog();

        //    if (processError != null)
        //    {
        //        System.Windows.Forms.MessageBox.Show(
        //            "Error during highlighting: " + processError.Message,
        //            "Error",
        //            System.Windows.Forms.MessageBoxButtons.OK,
        //            System.Windows.Forms.MessageBoxIcon.Error
        //        );
        //    }
        //}


        //public void HighlightAllAbbreviations()
        //{
        //    if (!isAbbreviationEnabled) return;

        //    Word.Document doc = this.Application.ActiveDocument;
        //    this.Application.ScreenUpdating = false;
        //    this.Application.DisplayStatusBar = false;

        //    try
        //    {
        //        if (!AbbreviationManager.IsAutoCorrectCacheInitialized())
        //            AbbreviationManager.InitializeAutoCorrectCache(this.Application.AutoCorrect);

        //        var phrases = AbbreviationManager.GetAllPhrases()
        //            .OrderByDescending(p => p.Length) // longest first
        //            .ToList();

        //        var words = doc.Words;
        //        int totalWords = words.Count;

        //        for (int i = 1; i <= totalWords; i++)
        //        {
        //            Word.Range currentRange = words[i];
        //            foreach (var phrase in phrases)
        //            {
        //                string text = currentRange.Text.Trim();
        //                if (string.Equals(text, phrase, StringComparison.OrdinalIgnoreCase))
        //                {
        //                    currentRange.HighlightColorIndex = Word.WdColorIndex.wdYellow;
        //                    currentRange.Font.Color = Word.WdColor.wdColorRed;
        //                    break; 
        //                }
        //            }
        //        }
        //    }
        //    finally
        //    {
        //        this.Application.ScreenUpdating = true;
        //        this.Application.DisplayStatusBar = true;
        //    }
        //}




        /// <summary>
        /// Event: User typed text in the pane input box.
        /// </summary>
        public void SuggestionPaneControl_OnTextChanged(string inputText)
        {
            try
            {
                frozeSuggestions = false;

                // Get the SuggestionPaneControl for the active window
                var currentControl = EnsureTaskPaneVisible(this.Application.ActiveWindow, "onTextChanges");
                if (currentControl == null) return;

                List<(string Word, string Replacement)> matches = new List<(string Word, string Replacement)>();

                if (currentControl.CurrentMode == SuggestionPaneControl.Mode.Reverse)
                {
                    if (!string.IsNullOrEmpty(inputText))
                    {
                        matches = AbbreviationManager.GetAllPhrases()
                            .Select(abbrev =>
                            {
                                string full = AbbreviationManager.GetAbbreviation(abbrev);
                                return (Word: abbrev, Replacement: full);
                            })
                            .Where(p => !string.IsNullOrEmpty(p.Replacement) &&
                                        p.Replacement.StartsWith(inputText, StringComparison.InvariantCultureIgnoreCase))
                            .ToList();
                    }
                }
                else
                {
                    // Abbreviation and Dictionary mode
                    matches = trie.GetWordsWithPrefix(inputText.ToLowerInvariant())
                        .Select(p => (Word: p, Replacement: AbbreviationManager.GetAbbreviation(p)))
                        .ToList();
                }

                var mode = currentControl.CurrentMode;
                currentControl.ShowSuggestions(matches, mode);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error in OnTextChanged: " + ex.Message + "\n" + ex.StackTrace,
                    "Exception", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }




        /// <summary>
        /// Event: User accepted a suggestion.
        /// </summary>
        //private void SuggestionPaneControl_OnSuggestionAccepted(string inputText, string fullForm)
        //{
        //    isReplacing = true;
        //    try
        //    {
        //        Word.Selection sel = this.Application.Selection;
        //        if (sel == null) return;

        //        Word.Range sentenceRange = sel.Range.Sentences.First; // Or use Paragraphs.First
        //        string sentenceText = sentenceRange.Text;

        //        // Find exact match of the input phrase in the sentence
        //        int index = sentenceText.IndexOf(inputText, StringComparison.InvariantCultureIgnoreCase);
        //        if (index >= 0)
        //        {
        //            Word.Range matchRange = sentenceRange.Duplicate;
        //            matchRange.Start = sentenceRange.Start + index;
        //            matchRange.End = matchRange.Start + inputText.Length;

        //            matchRange.Text = fullForm + " ";
        //            sel.SetRange(matchRange.End, matchRange.End);

        //            // Optionally add to AutoCorrect
        //            var autoCorrect = this.Application.AutoCorrect;
        //            if (!autoCorrect.Entries.Cast<Word.AutoCorrectEntry>().Any(entry => entry.Name == inputText))
        //            {
        //                autoCorrect.Entries.Add(inputText, fullForm);
        //            }
        //        }
        //        else
        //        {
        //            MessageBox.Show($"Phrase not found in sentence.\nInput = '{inputText}'\nSentence = '{sentenceText}'");
        //        }
        //    }
        //    finally
        //    {
        //        isReplacing = false;
        //    }
        //}

        //private void SuggestionPaneControl_OnSuggestionAccepted(string shortForm, string abbreviation)
        //{
        //    try
        //    {
        //        isReplacing = true;

        //        Word.Selection sel = this.Application.Selection;
        //        if (sel == null || sel.Range == null) return;

        //        if (string.IsNullOrEmpty(abbreviation)) return;

        //        Word.Range replaceRange = sel.Range.Duplicate;

        //        // Go back by number of words in short form
        //        int wordCount = shortForm.Split(' ').Length;
        //        replaceRange.MoveStart(Word.WdUnits.wdWord, -wordCount);

        //        string rangeText = replaceRange.Text?.Trim();

        //        if (!string.IsNullOrEmpty(rangeText) &&
        //            string.Equals(rangeText, shortForm, StringComparison.InvariantCultureIgnoreCase))
        //        {
        //            replaceRange.Text = abbreviation + " ";
        //            sel.SetRange(replaceRange.End, replaceRange.End);
        //        }
        //        else
        //        {
        //            // fallback by character length
        //            Word.Range fallbackRange = sel.Range.Duplicate;
        //            fallbackRange.MoveStart(Word.WdUnits.wdCharacter, -shortForm.Length);
        //            fallbackRange.Text = abbreviation + " ";
        //            sel.SetRange(fallbackRange.End, fallbackRange.End);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        System.Diagnostics.Debug.WriteLine("Replacement Error: " + ex.Message);
        //    }
        //    finally
        //    {
        //        isReplacing = false;
        //    }
        //}

        private void SuggestionPaneControl_OnSuggestionAccepted(string shortForm, string abbreviation)
        {
            return;
            //try
            //{
            //    isReplacing = true;

            //    Word.Selection sel = this.Application.Selection;
            //    if (sel == null || sel.Range == null) return;

            //    if (string.IsNullOrEmpty(abbreviation)) return;

            //    int wordCount = shortForm.Split(' ').Length;

            //    // Duplicate selection and move back by word count
            //    Word.Range candidateRange = sel.Range.Duplicate;
            //    candidateRange.MoveStart(Word.WdUnits.wdWord, -wordCount);

            //    string candidateText = candidateRange.Text?.Trim();

            //    // ✅ Only replace if the last words match exactly
            //    if (string.Equals(candidateText, shortForm, StringComparison.InvariantCultureIgnoreCase))
            //    {
            //        candidateRange.Text = abbreviation + " ";
            //        sel.SetRange(candidateRange.End, candidateRange.End);
            //    }
            //    else
            //    {
            //        // ❌ Not matching, so don't replace whole phrase — fallback to replacing only current word
            //        Word.Range fallbackRange = sel.Range.Duplicate;
            //        fallbackRange.MoveStart(Word.WdUnits.wdWord, -1);
            //        fallbackRange.Text = abbreviation + " ";
            //        sel.SetRange(fallbackRange.End, fallbackRange.End);
            //    }
            //}
            //catch (Exception ex)
            //{
            //    System.Diagnostics.Debug.WriteLine("Replacement Error: " + ex.Message);
            //}
            //finally
            //{
            //    isReplacing = false;
            //}
        }



        private void EnsurePhraseCache()
        {
            if (_phraseCache == null)
            {
                _phraseCache = AbbreviationManager.GetAllPhrases()
                    .Select(p => (Word: p, Replacement: GetFullFormFor(p)))
                    .ToList();
            }
        }

        private void TypingTimer_Tick(object sender, EventArgs e)
        {
            if (isReplacing || isProgrammaticChange)
                return;

            try
            {
                if (this.Application.Documents.Count == 0)
                    return;

                var sel = this.Application.Selection;
                if (sel?.Range == null)
                    return;

                int currentCharCount = sel.Document.Content.Characters.Count;
                int selStart = sel.Range.Start;

                bool charChanged =
                    lastDocumentCharCount != -1 &&
                    currentCharCount != lastDocumentCharCount;

                bool cursorMoved =
                    lastSelectionStart != -1 &&
                    selStart != lastSelectionStart;

                if (charChanged || cursorMoved)
                {
                    isTypingDetected = true;
                    debounceTimer.Stop();
                    debounceTimer.Start();
                }
                else
                {
                    isTypingDetected = false;
                }

                lastDocumentCharCount = currentCharCount;
                lastSelectionStart = selStart;
            }
            catch
            {
                // Never crash Word
            }
        }






        private void DebounceTimer_Tick(object sender, EventArgs e)
        {
            debounceTimer.Stop();



            if (this.Application.Documents.Count == 0 || !isAbbreviationEnabled)
                return;


            var selec = this.Application.Selection;
            if (selec?.Range == null)
                return;


            if (selec.Range.StoryType != Word.WdStoryType.wdMainTextStory)
                return;


            var window = this.Application.ActiveWindow;
            if (window == null) return;


            SuggestionPaneControl currentControl = null;

            if (!taskPanes.ContainsKey(window))
            {
                currentControl = EnsureTaskPaneVisible(window, "Debouncer");
            }
            else
            {
                taskPanes.TryGetValue(window, out var pane);
                currentControl = pane?.Control as SuggestionPaneControl;
            }

            try
            {
                if (currentControl == null)
                    return;

                Word.Selection sel = this.Application.Selection;
                if (sel?.Range == null) return;

                Word.Range selRange = sel.Range.Duplicate;

                int wordsChecked = 0;
                Word.Range testRange = selRange.Duplicate;

                while (wordsChecked < maxPhraseLength)
                {
                    if (testRange.Start == 0) break;

                    testRange.MoveStart(Word.WdUnits.wdWord, -1);
                    string candidate = testRange.Text.Trim();
                    if (string.IsNullOrEmpty(candidate))
                    {
                        wordsChecked++;
                        continue;
                    }

                    if (candidate.Contains("\r") || candidate.Contains("\n"))
                    {
                        testRange.MoveStart(Word.WdUnits.wdWord, 1);
                        break;
                    }

                    if (candidate.Length < 3)
                    {
                        wordsChecked++;
                        continue;
                    }

                    // Check last replaced word logic
                    if (!string.IsNullOrEmpty(lastReplacedShortForm) && !string.IsNullOrEmpty(lastReplacedFullForm))
                    {
                        if (string.Equals(candidate, lastReplacedShortForm, StringComparison.InvariantCultureIgnoreCase) &&
                            !string.Equals(candidate, lastReplacedFullForm, StringComparison.InvariantCultureIgnoreCase))
                        {
                            lastUndoneWord = lastReplacedShortForm;
                        }
                    }

                    if (!string.IsNullOrEmpty(lastUndoneWord) &&
                        string.Equals(candidate, lastUndoneWord, StringComparison.InvariantCultureIgnoreCase))
                    {
                        return;
                    }

                    List<(string Word, string Replacement)> matchesAbbrev = null;
                    List<(string Word, string Replacement)> matchesReverse = null;

                    // Search in Abbreviation mode (phrase -> abbreviation)
                    matchesAbbrev = trie.GetWordsWithPrefix(candidate.ToLowerInvariant())
                        .Select(p => (Word: p, Replacement: AbbreviationManager.GetAbbreviation(p)))
                        .ToList();

                    // Search in Reverse mode (abbreviation -> phrase)
                    matchesReverse = AbbreviationManager.GetAllPhrases()
                        .Select(p => (Word: p, Replacement: AbbreviationManager.GetAbbreviation(p)))
                        .Where(p => !string.IsNullOrEmpty(p.Replacement) &&
                                    p.Replacement.StartsWith(candidate, StringComparison.InvariantCultureIgnoreCase))
                        .ToList();

                    // Check if we have matches in either mode
                    bool hasAbbrevMatches = matchesAbbrev != null && matchesAbbrev.Count > 0;
                    bool hasReverseMatches = matchesReverse != null && matchesReverse.Count > 0;

                    if (!hasAbbrevMatches && !hasReverseMatches)
                    {
                        wordsChecked++;
                        continue;
                    }

                    // KEY CHANGE: Show task pane when matches are found in EITHER mode
                    //if (taskPanes.TryGetValue(window, out var taskPane))
                    //{
                    //    if (!taskPane.Visible)
                    //    {
                    //        taskPane.Visible = true;
                    //        // Remove from userClosedTaskPanes since we're reopening it due to matches
                    //        userClosedTaskPanes.Remove(window);
                    //    }
                    //}

                    // Show suggestions based on current mode
                    currentControl.SetInputText(candidate);
                    var mode = currentControl.CurrentMode;

                    // Display matches based on which mode the pane is in
                    if (mode == SuggestionPaneControl.Mode.Abbreviation)
                    {
                        // Normal mode: Show abbreviation suggestions from Word document typing
                        currentControl.ShowSuggestions(matchesAbbrev, mode); 
                    }
                    else if (mode == SuggestionPaneControl.Mode.Reverse)
                    {
                        // Reverse mode: DO NOT show suggestions from Word document typing
                        // Reverse suggestions should only come from input box typing
                        // Keep existing suggestions in reverse tab, don't update them
                        return; // Exit without updating suggestions
                    }

                    return;
                }

                lastUndoneWord = null;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Error in DebounceTimer_Tick: " + ex.Message + "\n" + ex.StackTrace,
                    "Exception",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error
                );
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


        public class TrieNode
        {
            public Dictionary<char, TrieNode> Children { get; } = new Dictionary<char, TrieNode>();
            public List<string> Words { get; } = new List<string>();
        }

        public class Trie
        {
            private readonly TrieNode root = new TrieNode();

            public void Insert(string word)
            {
                var node = root;
                foreach (char c in word)
                {
                    if (!node.Children.ContainsKey(c))
                    {
                        node.Children[c] = new TrieNode();
                    }
                    node = node.Children[c];
                    node.Words.Add(word);
                }
            }

            public List<string> GetWordsWithPrefix(string prefix)
            {
                var node = root;
                foreach (char c in prefix)
                {
                    if (!node.Children.ContainsKey(c))
                    {
                        return new List<string>();
                    }
                    node = node.Children[c];
                }
                return node.Words.Distinct().ToList();
            }
        }





        /// <summary>
        /// Public method to force JSSD ribbon tab visibility - can be called from outside
        /// </summary>
        public void ForceJSSDRibbonVisible()
        {
            try
            {
                var ribbonTab = Globals.Ribbons?.AbbreviationRibbon;
                if (ribbonTab != null)
                {
                    ribbonTab.ForceJSSDTabVisible();
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("JSSD ribbon tab not found. Please restart Word.", "Ribbon Not Found", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error forcing JSSD tab visibility: " + ex.Message, "Ribbon Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
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
