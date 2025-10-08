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
            var app = Globals.ThisAddIn.Application;

            string helpPath = ExtractTemplateToLocal("AbbreviationWordAddin.Help.Help.docx", "Help.docx");

            Microsoft.Office.Interop.Word.Document helpDoc = app.Documents.Open(
                FileName: helpPath,
                ReadOnly: true,   
                Visible: true
            );

            if (!helpDoc.ProtectionType.HasFlag(Microsoft.Office.Interop.Word.WdProtectionType.wdAllowOnlyReading))
            {
                helpDoc.Protect(
                    Type: Microsoft.Office.Interop.Word.WdProtectionType.wdAllowOnlyReading,
                    NoReset: true,      
                    Password: "1234"        
                );
            }

            helpDoc.Saved = true;
        }




    }
}
