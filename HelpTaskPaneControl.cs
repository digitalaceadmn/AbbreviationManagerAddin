using Word = Microsoft.Office.Interop.Word;
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

            this.AutoScaleDimensions = new SizeF(8F, 16F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.BackColor = Color.White;
            this.Size = new Size(400, 600);

            var headerPanel = new Panel
            {
                BackColor = Color.FromArgb(0, 120, 215),
                Height = 60,
                Dock = DockStyle.Top
            };

            var titleLabel = new Label
            {
                Text = "📘 Help - Abbreviation Manager",
                Font = new Font("Segoe UI", 14F, FontStyle.Bold),
                ForeColor = Color.White,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill
            };

            headerPanel.Controls.Add(titleLabel);

            var contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(0)
            };

            var helpBrowser = new WebBrowser
            {
                Dock = DockStyle.Fill,
                AllowWebBrowserDrop = false,
                IsWebBrowserContextMenuEnabled = false,
                ScriptErrorsSuppressed = true
            };

            contentPanel.Controls.Add(helpBrowser);

            this.Controls.Add(contentPanel);
            this.Controls.Add(headerPanel);

            this.ResumeLayout(false);
        }

        private WebBrowser GetHelpBrowser()
        {
            foreach (Control c in this.Controls)
            {
                if (c is Panel panel)
                {
                    foreach (Control child in panel.Controls)
                    {
                        if (child is WebBrowser wb)
                            return wb;
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

        private void LoadWordHelpIntoBrowser(string docxPath, WebBrowser browser)
        {
            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                wordApp = Globals.ThisAddIn.Application;

                doc = wordApp.Documents.Open(
                    docxPath,
                    ReadOnly: true,
                    Visible: false
                );

                string tempHtmlPath = Path.Combine(
                    Path.GetTempPath(),
                    Guid.NewGuid().ToString("N") + ".html"
                );

                // DOCX → HTML (IMAGES SUPPORTED)
                doc.SaveAs2(
                    tempHtmlPath,
                    Word.WdSaveFormat.wdFormatFilteredHTML
                );

                doc.Close(false);
                doc = null;

                browser.Navigate(tempHtmlPath);
            }
            catch (Exception ex)
            {
                browser.DocumentText =
                    "<html><body style='font-family:Segoe UI'>" +
                    "<h3>Error loading help</h3><pre>" +
                    ex.Message +
                    "</pre></body></html>";
            }
            finally
            {
                if (doc != null)
                {
                    try { doc.Close(false); } catch { }
                }
            }
        }

        private void LoadHelpContent()
        {
            WebBrowser browser = GetHelpBrowser();

            if (browser == null)
            {
                MessageBox.Show("Help browser not found.");
                return;
            }

            string embeddedName = "AbbreviationWordAddin.Help.Help.docx";
            string templatePath = ExtractTemplateToLocal(embeddedName, "Help.docx");

            if (templatePath == null)
            {
                browser.DocumentText =
                    "<html><body>Help file could not be extracted.</body></html>";
                return;
            }

            LoadWordHelpIntoBrowser(templatePath, browser);
        }

        private void HelpTaskPaneControl_Load(object sender, EventArgs e)
        {
            LoadHelpContent();
        }
    }
}
