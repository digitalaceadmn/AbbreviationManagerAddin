using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace AbbreviationWordAddin
{
    public partial class HelpTaskPaneControl : UserControl
    {
        public HelpTaskPaneControl()
        {
            InitializeComponent();
            LoadHelpContent();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // Main container
            this.AutoScaleDimensions = new SizeF(8F, 16F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.BackColor = Color.White;
            this.Size = new Size(400, 600);

            // Header Panel
            var headerPanel = new Panel();
            headerPanel.BackColor = Color.FromArgb(0, 120, 215); // Office blue
            headerPanel.Height = 60;
            headerPanel.Dock = DockStyle.Top;

            var titleLabel = new Label();
            titleLabel.Text = "🔒 Help - Abbreviation Manager";
            titleLabel.Font = new Font("Segoe UI", 14F, FontStyle.Bold);
            titleLabel.ForeColor = Color.White;
            titleLabel.AutoSize = false;
            titleLabel.TextAlign = ContentAlignment.MiddleCenter;
            titleLabel.Dock = DockStyle.Fill;
            headerPanel.Controls.Add(titleLabel);

            // Content Panel with scroll
            var contentPanel = new Panel();
            contentPanel.Dock = DockStyle.Fill;
            contentPanel.AutoScroll = true;
            contentPanel.Padding = new Padding(15);

            // Main content RichTextBox (read-only)
            var helpContent = new RichTextBox();
            helpContent.ReadOnly = true;
            helpContent.BorderStyle = BorderStyle.None;
            helpContent.BackColor = Color.White;
            helpContent.Dock = DockStyle.Fill;
            helpContent.Font = new Font("Segoe UI", 10F);
            helpContent.SelectionIndent = 10;
            helpContent.SelectionHangingIndent = -10;

            contentPanel.Controls.Add(helpContent);

            // Add controls to main form
            this.Controls.Add(contentPanel);
            this.Controls.Add(headerPanel);

            this.ResumeLayout(false);
        }

        private RichTextBox GetHelpRichTextBox()
        {
            foreach (Control c in this.Controls)
            {
                if (c is Panel panel)
                {
                    foreach (Control child in panel.Controls)
                    {
                        if (child is RichTextBox rtb)
                            return rtb;
                    }
                }
            }
            return null;
        }


        public static string ExtractTemplateToLocal(string embeddedName, string outputName)
        {
            string outputDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "AbbreviationWordAddin",
                "Help"
            );

            Directory.CreateDirectory(outputDir);

            string fullPath = Path.Combine(outputDir, outputName);

            Assembly asm = Assembly.GetExecutingAssembly();

            // IMPORTANT:
            // You MUST include full namespace + folder
            // e.g. "AbbreviationWordAddin.Help.Help.docx"
            Stream stream = asm.GetManifestResourceStream(embeddedName);

            if (stream == null)
            {
                MessageBox.Show(
                    "Embedded help file NOT found:\n" + embeddedName +
                    "\n\nAvailable resources:\n" +
                    string.Join("\n", asm.GetManifestResourceNames())
                );
                return null;
            }

            using (stream)
            using (FileStream fs = new FileStream(fullPath, FileMode.Create, FileAccess.Write))
            {
                stream.CopyTo(fs);
            }

            return fullPath;
        }


        private string LoadWordHelpText(string filePath)
        {
            if (!File.Exists(filePath))
                return "Help file not found.\n\nExpected: " + filePath;

            try
            {
                var wordApp = Globals.ThisAddIn.Application;
                var doc = wordApp.Documents.Open(filePath, ReadOnly: true, Visible: false);

                string text = doc.Content.Text;

                doc.Close(false);

                return text;
            }
            catch (Exception ex)
            {
                return "Error loading help file:\n" + ex.Message;
            }
        }


        private void LoadHelpContent()
        {
            RichTextBox helpContent = GetHelpRichTextBox();

            if (helpContent == null)
            {
                MessageBox.Show("Error: Help RichTextBox not found.");
                return;
            }

            // Correct embedded resource name:
            // Example: "AbbreviationWordAddin.Help.Help.docx"
            string embeddedName = "AbbreviationWordAddin.Help.Help.docx";

            string templatePath = ExtractTemplateToLocal(embeddedName, "Help.docx");

            if (templatePath == null)
            {
                helpContent.Text = "Help file could not be extracted.";
                return;
            }

            string helpText = LoadWordHelpText(templatePath);

            helpContent.Clear();
            helpContent.Text = helpText ?? "Unable to load help content.";

            helpContent.ReadOnly = true;
            helpContent.TabStop = false;
        }



        private string CreateHelpContentRTF()
        {
            return @"{\rtf1\ansi\deff0 {\fonttbl {\f0 Segoe UI;} {\f1 Segoe UI;} {\f2 Segoe UI;}}
{\colortbl;\red0\green120\blue215;\red0\green0\blue0;\red255\green255\blue255;\red40\green40\blue40;}

\f0\fs28\cf1\b ABBREVIATION MANAGER - USER GUIDE\b0\fs20\cf2\par\par

\f1\fs24\cf1\b OVERVIEW\b0\fs20\cf2\par
This add-in helps you manage and use abbreviations in Microsoft Word documents efficiently.\par\par

\f1\fs22\cf1\b KEY FEATURES:\b0\fs20\cf2\par
\bullet\tab Forward Abbreviations: Type abbreviations to expand to full text\par
\bullet\tab Reverse Abbreviations: Select from list to insert abbreviations\par
\bullet\tab Dictionary View: Browse all available abbreviations\par
\bullet\tab Bulk Operations: Replace or highlight all abbreviations at once\par\par

\f1\fs22\cf1\b HOW TO USE:\b0\fs20\cf2\par\par

\f2\fs20\cf4\b 1. ENABLE ABBREVIATIONS\b0\cf2\par
\bullet\tab Click 'Enable' in the Abbreviation ribbon tab\par
\bullet\tab The suggestion pane will appear automatically\par\par

\f2\fs20\cf4\b 2. FORWARD ABBREVIATIONS (Normal Mode)\b0\cf2\par
\bullet\tab Switch to 'Abbreviation' tab in the suggestion pane\par
\bullet\tab Start typing in your document\par
\bullet\tab Suggestions will appear as you type\par
\bullet\tab Double-click any suggestion to insert it\par\par

\f2\fs20\cf4\b 3. REVERSE ABBREVIATIONS (Reverse Mode)\b0\cf2\par
\bullet\tab Switch to 'Reverse' tab in the suggestion pane\par
\bullet\tab Browse the list of available full forms\par
\bullet\tab Double-click any item to insert its abbreviation\par
\bullet\tab \cf1\b NOTE: Typing does NOT generate suggestions in reverse mode\b0\cf2\par\par

\f2\fs20\cf4\b 4. DICTIONARY VIEW\b0\cf2\par
\bullet\tab Switch to 'Dictionary' tab to view all abbreviations\par
\bullet\tab Alphabetically sorted for easy browsing\par
\bullet\tab Shows abbreviation and full form pairs\par\par

\f2\fs20\cf4\b 5. BULK OPERATIONS\b0\cf2\par
\bullet\tab 'Replace All': Find and replace all abbreviations in document\par
\bullet\tab 'Highlight All': Highlight all potential abbreviations\par
\bullet\tab 'Highlight Like': Advanced pattern highlighting\par\par

\f1\fs22\cf1\b RIBBON BUTTONS:\b0\fs20\cf2\par
\bullet\tab \b Enable/Disable\b0: Toggle abbreviation functionality\par
\bullet\tab \b Replace All\b0: Batch replace abbreviations in document\par
\bullet\tab \b Highlight All\b0: Highlight abbreviations for review\par
\bullet\tab \b Show Suggestions\b0: Display/hide the suggestion pane\par
\bullet\tab \b Templates\b0: Access document templates\par
\bullet\tab \b Help\b0: Show this help information\par\par

\f1\fs22\cf1\b TIPS & BEST PRACTICES:\b0\fs20\cf2\par
\bullet\tab Keep the suggestion pane open while working\par
\bullet\tab Use 'Dictionary' tab to familiarize yourself with available abbreviations\par
\bullet\tab For reverse abbreviations, always select from the list rather than typing\par
\bullet\tab Use 'Replace All' for quick document processing\par
\bullet\tab Use 'Highlight All' to review abbreviations before replacing\par\par

\f1\fs22\cf1\b TROUBLESHOOTING:\b0\fs20\cf2\par
\bullet\tab If suggestions don't appear, try clicking 'Enable' again\par
\bullet\tab If the pane disappears, click 'Show Suggestions'\par
\bullet\tab For reverse mode, ensure you're selecting from the list, not typing\par
\bullet\tab Restart Word if abbreviations stop working\par\par

\f1\fs22\cf1\b SECURITY NOTE:\b0\fs20\cf2\par
This help content is displayed in a read-only task pane and cannot be edited.\par
All abbreviation data is securely managed by the add-in.\par\par

\cf1\b For additional support, contact your system administrator.\b0\cf2\par
}";
        }

        private void HelpTaskPaneControl_Load(object sender, EventArgs e)
        {
            // Ensure content is loaded when control loads
            LoadHelpContent();
        }
    }
}