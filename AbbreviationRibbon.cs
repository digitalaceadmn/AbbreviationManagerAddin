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
            try
            {
                // Manually set the JSSD tab label to ensure it's visible in new Word versions
                this.JSSD.Label = "JSSD";
                this.JSSD.Visible = false;
                
                // Ensure all groups are visible
                this.group1.Visible = true; // Replace / Highlight
                this.group2.Visible = false; // Enable / Disable  
                this.group3.Visible = true; // Template / Show Suggestions
                this.group4.Visible = true; // Help
                
                // Properties set - no explicit refresh needed
            }
            catch (Exception)
            {
                // Silent catch - don't interrupt normal loading
            }
            
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
            Globals.ThisAddIn.IgnoredAbbreviations.Clear();

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
                button.Label = "List All";  
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

        /// <summary>
        /// Public method to force refresh and ensure JSSD tab visibility
        /// Call this method if the JSSD tab is not visible in newer Word versions
        /// </summary>
        public void ForceJSSDTabVisible()
        {
            try
            {
                // Force set the JSSD tab properties
                if (this.JSSD != null)
                {
                    this.JSSD.Label = "JSSD";
                    this.JSSD.Visible = false;
                }
                
                // Make sure all groups are visible
                if (this.group1 != null) this.group1.Visible = true;
                if (this.group2 != null) this.group2.Visible = true;
                if (this.group3 != null) this.group3.Visible = true;
                if (this.group4 != null) this.group4.Visible = true;
            }
            catch (Exception)
            {
                // Silent catch - don't show error messages during normal operation
            }
        }
        
        public void ForceJSSDTabVisibleWithMessage()
        {
            try
            {
                ForceJSSDTabVisible();
                System.Windows.Forms.MessageBox.Show("JSSD tab has been refreshed and should now be visible.", "Ribbon Refresh", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error refreshing JSSD tab: " + ex.Message, "Ribbon Refresh Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
        
        public void TestRibbonAccess()
        {
            try
            {
                // Simple test to verify ribbon is accessible
                string tabName = this.JSSD?.Label ?? "Not Found";
                bool tabVisible = this.JSSD?.Visible ?? false;
                
                System.Windows.Forms.MessageBox.Show(
                    $"JSSD Tab Status:\nLabel: {tabName}\nVisible: {tabVisible}\n\nIf visible is False, the tab may be hidden by Word.", 
                    "Ribbon Status Test", 
                    System.Windows.Forms.MessageBoxButtons.OK, 
                    System.Windows.Forms.MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error testing ribbon: " + ex.Message, "Test Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

    }
}
