using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace AbbreviationWordAddin
{
    partial class AbbreviationRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AbbreviationRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        public static class TemplatePaths
        {
            private static readonly string BaseDir = AppDomain.CurrentDomain.BaseDirectory;

            public static string GeneralFormat => Path.Combine(BaseDir, "Templates", "GeneralFormat.dotx");
            public static string ServiceLetterFormat => Path.Combine(BaseDir, "Templates", "ServiceLetterFormat.dotx");
            public static string SalaryCertificateFormat => Path.Combine(BaseDir, "Templates", "SalaryCertificateFormat.dotx");
            public static string ExperienceCertificateFormat => Path.Combine(BaseDir, "Templates", "ExperienceCertificateFormat.dotx");
            public static string OfferLetterFormat => Path.Combine(BaseDir, "Templates", "OfferLetterFormat.dotx");
            public static string AppointmentLetterFormat => Path.Combine(BaseDir, "Templates", "AppointmentLetterFormat.dotx");
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.AppxC = this.Factory.CreateRibbonTab();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnEnable = this.Factory.CreateRibbonButton();
            this.btnDisable = this.Factory.CreateRibbonButton();
            this.btnReplaceAll = this.Factory.CreateRibbonButton();
            this.btnHighlightAll = this.Factory.CreateRibbonButton();
            this.menuTemplates = this.Factory.CreateRibbonMenu();
            this.btnGeneralFormat = this.Factory.CreateRibbonButton();
            this.btnServiceLetterFormat = this.Factory.CreateRibbonButton();
            this.btnSalaryCertificateFormat = this.Factory.CreateRibbonButton();
            this.btnExperienceCertificateFormat = this.Factory.CreateRibbonButton();
            this.btnOfferLetterFormat = this.Factory.CreateRibbonButton();
            this.btnAppointmentLetterFormat = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.AppxC.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // AppxC
            // 
            this.AppxC.Groups.Add(this.group2);
            this.AppxC.Label = "Appx-C";
            this.AppxC.Name = "AppxC";
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnEnable);
            this.group2.Items.Add(this.btnDisable);
            this.group2.Items.Add(this.btnReplaceAll);
            this.group2.Items.Add(this.btnHighlightAll);
            this.group2.Items.Add(this.menuTemplates);
            this.group2.Label = "Abbreviation Tools";
            this.group2.Name = "group2";
            // 
            // btnEnable
            // 
            this.btnEnable.Label = "Enable Abbreviation";
            this.btnEnable.Name = "btnEnable";
            this.btnEnable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEnable_Click);
            // 
            // btnDisable
            // 
            this.btnDisable.Label = "Disable Abbreviation";
            this.btnDisable.Name = "btnDisable";
            this.btnDisable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDisable_Click);
            // 
            // btnReplaceAll
            // 
            this.btnReplaceAll.Label = "Replace All";
            this.btnReplaceAll.Name = "btnReplaceAll";
            this.btnReplaceAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReplaceAll_Click);
            // 
            // btnHighlightAll
            // 
            this.btnHighlightAll.Label = "Highlight All";
            this.btnHighlightAll.Name = "btnHighlightAll";
            this.btnHighlightAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnHighlightAll_Click);
            // 
            // menuTemplates
            // 
            this.menuTemplates.Items.Add(this.btnGeneralFormat);
            this.menuTemplates.Items.Add(this.btnServiceLetterFormat);
            this.menuTemplates.Items.Add(this.btnSalaryCertificateFormat);
            this.menuTemplates.Items.Add(this.btnExperienceCertificateFormat);
            this.menuTemplates.Items.Add(this.btnOfferLetterFormat);
            this.menuTemplates.Items.Add(this.btnAppointmentLetterFormat);
            this.menuTemplates.Label = "Template";
            this.menuTemplates.Name = "menuTemplates";
            // 
            // btnGeneralFormat
            // 
            this.btnGeneralFormat.Label = "General Format";
            this.btnGeneralFormat.Name = "btnGeneralFormat";
            this.btnGeneralFormat.ShowImage = true;
            this.btnGeneralFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGeneralFormat_Click);
            // 
            // btnServiceLetterFormat
            // 
            this.btnServiceLetterFormat.Label = "Service Letter Format";
            this.btnServiceLetterFormat.Name = "btnServiceLetterFormat";
            this.btnServiceLetterFormat.ShowImage = true;
            this.btnServiceLetterFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnServiceLetterFormat_Click);
            // 
            // btnSalaryCertificateFormat
            // 
            this.btnSalaryCertificateFormat.Label = "Salary Certificate Format";
            this.btnSalaryCertificateFormat.Name = "btnSalaryCertificateFormat";
            this.btnSalaryCertificateFormat.ShowImage = true;
            this.btnSalaryCertificateFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSalaryCertificateFormat_Click);
            // 
            // btnExperienceCertificateFormat
            // 
            this.btnExperienceCertificateFormat.Label = "Experience Certificate Format";
            this.btnExperienceCertificateFormat.Name = "btnExperienceCertificateFormat";
            this.btnExperienceCertificateFormat.ShowImage = true;
            this.btnExperienceCertificateFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExperienceCertificateFormat_Click);
            // 
            // btnOfferLetterFormat
            // 
            this.btnOfferLetterFormat.Label = "Offer Letter Format";
            this.btnOfferLetterFormat.Name = "btnOfferLetterFormat";
            this.btnOfferLetterFormat.ShowImage = true;
            this.btnOfferLetterFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOfferLetterFormat_Click);
            // 
            // btnAppointmentLetterFormat
            // 
            this.btnAppointmentLetterFormat.Label = "Appointment Letter Format";
            this.btnAppointmentLetterFormat.Name = "btnAppointmentLetterFormat";
            this.btnAppointmentLetterFormat.ShowImage = true;
            this.btnAppointmentLetterFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAppointmentLetterFormat_Click);
            // 
            // AbbreviationRibbon
            // 
            this.Name = "AbbreviationRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.AppxC);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.AbbreviationRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.AppxC.ResumeLayout(false);
            this.AppxC.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }


        private void btnGeneralFormat_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.GeneralFormat.docx", "GeneralFormat.docx");

            if (!File.Exists(templatePath))
            {
                MessageBox.Show("Template not found: " + templatePath);
                return;
            }

            try
            {
                var tempDoc = Globals.ThisAddIn.Application.Documents.Open(templatePath, Visible: false);

                tempDoc.Content.WholeStory();
                tempDoc.Content.Copy();

                tempDoc.Close(false);

                System.Threading.Thread.Sleep(200); // <-- critical in some environments

                var currentDoc = Globals.ThisAddIn.Application.ActiveDocument;
                currentDoc.Activate();

                Globals.ThisAddIn.Application.Selection.HomeKey(WdUnits.wdStory);

                Globals.ThisAddIn.Application.Selection.Paste();

                MessageBox.Show("Template inserted successfully!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error inserting template: " + ex.Message);
            }
        }


        private void btnServiceLetterFormat_Click(object sender, RibbonControlEventArgs e)
        {
            OpenTemplate("ServiceLetterFormat.dotx");
        }

        private void btnSalaryCertificateFormat_Click(object sender, RibbonControlEventArgs e)
        {
            OpenTemplate("ServiceLetterFormat.dotx");
        }

        private void btnExperienceCertificateFormat_Click(object sender, RibbonControlEventArgs e)
        {
            OpenTemplate("GeneralFormat.dotx");
        }

        private void btnOfferLetterFormat_Click(object sender, RibbonControlEventArgs e)
        {
            OpenTemplate("GeneralFormat.dotx");
        }

        private void btnAppointmentLetterFormat_Click(object sender, RibbonControlEventArgs e)
        {
            OpenTemplate("GeneralFormat.dotx");
        }

        public static string ExtractTemplateToLocal(string resourceName, string outputFileName)
        {
            string outputDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "AbbreviationWordAddin", "Templates");
            Directory.CreateDirectory(outputDir);

            string fullPath = Path.Combine(outputDir, outputFileName);

            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
            using (FileStream fileStream = new FileStream(fullPath, FileMode.Create))
            {
                stream.CopyTo(fileStream);
            }

            return fullPath;
        }

        private void OpenTemplate(string templateFileName)
        {
            // Get the folder where Abbreviations.xlsx is located
            string basePath = System.IO.Path.GetDirectoryName(
                System.IO.Path.Combine(
                    System.Windows.Forms.Application.StartupPath,
                    "Abbreviations.xlsx"
                )
            );

            System.Windows.Forms.MessageBox.Show(
                "Base File path" + basePath,
                            "Abbreviation Loading status",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Information
                        );

            string templatePath = System.IO.Path.Combine(basePath, templateFileName);

            var wordApp = Globals.ThisAddIn.Application;
            wordApp.Documents.Add(Template: templatePath);
        }

        private void InsertTemplate(string templateName)
        {
            try
            {
                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", templateName);
                if (!File.Exists(templatePath))
                {
                    MessageBox.Show("Template not found: " + templatePath);
                    return;
                }

                // Open template document invisibly
                var tempDoc = Globals.ThisAddIn.Application.Documents.Open(templatePath, Visible: false);

                // Copy its content
                tempDoc.Content.Copy();
                tempDoc.Close(false);

                // Paste into current document
                Globals.ThisAddIn.Application.Selection.Paste();

                MessageBox.Show("Template inserted successfully!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to insert template: " + ex.Message);
            }
        }



        private void InsertTemplateFromEmbeddedResource(string resourceName)
        {
            try
            {
                var wordApp = Globals.ThisAddIn.Application;
                Word.Document activeDoc = wordApp.ActiveDocument;

                var assembly = Assembly.GetExecutingAssembly();

                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream == null)
                    {
                        throw new Exception($"Resource not found: {resourceName}");
                    }

                    // Create a temp file to write the .dotx content
                    string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".dotx");
                    using (var fileStream = new FileStream(tempPath, FileMode.Create, FileAccess.Write))
                    {
                        stream.CopyTo(fileStream);
                    }

                    // Open the .dotx file invisibly
                    Word.Document tempDoc = wordApp.Documents.Open(tempPath, ReadOnly: true, Visible: false);
                    tempDoc.Content.Copy();
                    tempDoc.Close(SaveChanges: false);

                    wordApp.Selection.Paste();

                    System.Windows.Forms.MessageBox.Show("Template inserted successfully from embedded resource.");
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Error inserting embedded template: {ex.Message}",
                    "Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error
                );
            }
        }


        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab AppxC;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEnable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDisable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReplaceAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHighlightAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuTemplates;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGeneralFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnServiceLetterFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSalaryCertificateFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExperienceCertificateFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOfferLetterFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAppointmentLetterFormat;
    }

    partial class ThisRibbonCollection
    {
        internal AbbreviationRibbon AbbreviationRibbon
        {
            get { return this.GetRibbon<AbbreviationRibbon>(); }
        }
    }
}
