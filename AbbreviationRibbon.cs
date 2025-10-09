using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace AbbreviationWordAddin
{
    public partial class AbbreviationRibbon
    {
       
        private void AbbreviationRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            var application = Globals.ThisAddIn.Application;
            AutoCorrect autoCorrect = application.AutoCorrect;
            if (autoCorrect.ReplaceText)
            {
                AbbreviationManager.InitializeAutoCorrectCache(autoCorrect);

                Globals.ThisAddIn.ToggleAbbreviationReplacement(true);
                btnEnable.Enabled = false;  
                btnDisable.Enabled = true; 
                btnEnable.Label = "Enabled Abbreviations"; 
                btnDisable.Label = "Disable Abbreviations";
            }
            else
            {
                Globals.ThisAddIn.ToggleAbbreviationReplacement(false);
                btnEnable.Enabled = true;  
                btnDisable.Enabled = false; 
                btnDisable.Label = "Disabled"; 
                btnEnable.Label = "Enable"; 
            }
        }

       
        private void btnEnable_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                //Globals.ThisAddIn.SuggestionPaneControl?.SetSuggestionsVisible(true);
                Globals.ThisAddIn.isAbbreviationEnabled = true;
                var application = Globals.ThisAddIn.Application;
                AutoCorrect autoCorrect = application.AutoCorrect;

                AbbreviationManager.InitializeAutoCorrectCache(autoCorrect);
                
                Globals.ThisAddIn.ToggleAbbreviationReplacement(true);

                btnEnable.Enabled = false;  
                btnDisable.Enabled = true; 
                btnEnable.Label = "Enabled Abbreviations"; 
                btnDisable.Label = "Disable Abbreviations"; 

                autoCorrect.ReplaceText = true; 
                autoCorrect.CorrectCapsLock = true; 
                autoCorrect.CorrectSentenceCaps = true;
                autoCorrect.CorrectInitialCaps = true;
                autoCorrect.CorrectHangulAndAlphabet = true;
                autoCorrect.OtherCorrectionsAutoAdd = true;
            }
            catch (COMException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error while enabling abbreviator: " + ex.Message, "Abbreviator Enabling Error");
            }
        }

        private void btnDisable_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Globals.ThisAddIn.isAbbreviationEnabled = false;
                //Globals.ThisAddIn.SuggestionPaneControl?.SetSuggestionsVisible(false);

                var application = Globals.ThisAddIn.Application;
                AutoCorrect autoCorrect = application.AutoCorrect;

                AbbreviationManager.ClearAutoCorrectCache();
                
                Globals.ThisAddIn.ToggleAbbreviationReplacement(false);

                btnEnable.Enabled = true;  
                btnDisable.Enabled = false; 
                btnDisable.Label = "Disabled Abbreviations"; 
                btnEnable.Label = "Enable Abbreviations"; 

                autoCorrect.ReplaceText = false;
                autoCorrect.CorrectCapsLock = false;
                autoCorrect.CorrectSentenceCaps = false;
                autoCorrect.CorrectInitialCaps = false;
                autoCorrect.CorrectHangulAndAlphabet = false;
                autoCorrect.OtherCorrectionsAutoAdd = false;
            }
            catch (COMException ex)
            {
                System.Windows.Forms.MessageBox.Show("Error while disabling abbreviator: " + ex.Message, "Abbreviator Disabling Error");
            }
        }

        private async void btnReplaceAll_Click(object sender, RibbonControlEventArgs e)
        {
            var button = (RibbonButton)sender;
            button.Enabled = false; 
            button.Label = "Processing...";  

            try
            {
                if (!AbbreviationManager.IsAutoCorrectCacheInitialized())
                {
                    AbbreviationManager.InitializeAutoCorrectCache(Globals.ThisAddIn.Application.AutoCorrect);
                }

                await System.Threading.Tasks.Task.Run(() => Globals.ThisAddIn.ReplaceAllAbbreviations());
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error during abbreviation replace all: " + ex.Message, "Error");
            }
            finally
            {
                button.Label = "Replace All";  
                button.Enabled = true;  
            }
        }

        private async void btnHighlightAll_Click(object sender, RibbonControlEventArgs e)
        {
            var button = (RibbonButton)sender;
            button.Enabled = false; 
            button.Label = "Processing..."; 

            try
            {
                if (!AbbreviationManager.IsAutoCorrectCacheInitialized())
                {
                    AbbreviationManager.InitializeAutoCorrectCache(Globals.ThisAddIn.Application.AutoCorrect);
                }

                await System.Threading.Tasks.Task.Run(() => Globals.ThisAddIn.HighlightAllAbbreviations());
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error during highlighting abbreviation applicable phrases: " + ex.Message, "Error");
            }
            finally
            {
                button.Label = "Highlight All";  
                button.Enabled = true;  
            }
        }

        private void ShowSuggestions_Click(object sender, RibbonControlEventArgs e)
        {
            var window = Globals.ThisAddIn.Application.ActiveWindow;
            if (window == null) return;

            // If user had manually closed it before, allow reopen
            Globals.ThisAddIn.userClosedTaskPanes.Remove(window);

            var control = Globals.ThisAddIn.EnsureTaskPaneVisible(window, "Show Suggestion BTN");
            if (control == null) return;

            if (Globals.ThisAddIn.taskPanes.TryGetValue(window, out var pane))
            {
                pane.Visible = true;
            }
        }

        private async void highLightLike_Click(object sender, RibbonControlEventArgs e)
        {
            var button = (RibbonButton)sender;
            button.Enabled = false;  
            button.Label = "Processing...";  

            try
            {
                if (!AbbreviationManager.IsAutoCorrectCacheInitialized())
                {
                    AbbreviationManager.InitializeAutoCorrectCache(Globals.ThisAddIn.Application.AutoCorrect);
                }

                await System.Threading.Tasks.Task.Run(() => Globals.ThisAddIn.HighlightLike());
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error during highlighting abbreviation applicable phrases: " + ex.Message, "Error");
            }
            finally
            {
                button.Label = "Highlight Like";  
                button.Enabled = true; 
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var window = Globals.ThisAddIn.Application.ActiveWindow;
                if (window == null)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "No active Word window found. Please ensure Word is running and has a document open.",
                        "Help - Window Required",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information
                    );
                    return;
                }

                // Check if help task pane already exists for this window
                Microsoft.Office.Tools.CustomTaskPane helpTaskPane = null;
                
                // Look for existing help task pane
                foreach (Microsoft.Office.Tools.CustomTaskPane pane in Globals.ThisAddIn.CustomTaskPanes)
                {
                    if (pane.Title == "📖 Help - Abbreviation Manager" && pane.Window == window)
                    {
                        helpTaskPane = pane;
                        break;
                    }
                }

                // Create new help task pane if it doesn't exist
                if (helpTaskPane == null)
                {
                    var helpControl = new HelpTaskPaneControl();
                    helpTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(
                        helpControl,
                        "📖 Help - Abbreviation Manager",
                        window
                    );
                    
                    helpTaskPane.Width = 450;
                    helpTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                }

                // Show the help task pane
                helpTaskPane.Visible = true;
                
                // Show confirmation message
                System.Windows.Forms.MessageBox.Show(
                    "📖 HELP DISPLAYED IN TASK PANE 📖\n\n" +
                    "✓ Help content is now displayed in the task pane\n" +
                    "✓ Content is completely read-only and non-editable\n" +
                    "✓ Always accessible while working in Word\n" +
                    "✓ No separate document windows needed\n\n" +
                    "The help pane will remain open for easy reference.",
                    "Help Task Pane - Now Available",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information
                );
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Error displaying help task pane: " + ex.Message + "\n\n" +
                    "Please try again or restart Word if the problem persists.",
                    "Help Task Pane Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error
                );
            }
        }





    }
}
