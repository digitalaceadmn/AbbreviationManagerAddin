using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using static AbbreviationWordAddin.AbbreviationRibbon;
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
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnReplaceAll = this.Factory.CreateRibbonButton();
            this.btnHighlightAll = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.menuTemplates = this.Factory.CreateRibbonMenu();
            this.btnGeneralFormat = this.Factory.CreateRibbonButton();
            this.btnDoLetter = this.Factory.CreateRibbonButton();
            this.btnNotingSheet = this.Factory.CreateRibbonButton();
            this.btnSignalForm = this.Factory.CreateRibbonButton();
            this.btnStatementOfCase = this.Factory.CreateRibbonButton();
            this.btnStatementOfCaseDPM = this.Factory.CreateRibbonButton();
            this.btnAppxFormat = this.Factory.CreateRibbonButton();
            this.btnServicePaper = this.Factory.CreateRibbonButton();
            this.btnAgendaPts = this.Factory.CreateRibbonButton();
            this.btnOpNotes = this.Factory.CreateRibbonButton();
            this.btnMoM = this.Factory.CreateRibbonButton();
            this.btnEmailFormat = this.Factory.CreateRibbonButton();
            this.btnTourNotes = this.Factory.CreateRibbonButton();
            this.btnAppreciation = this.Factory.CreateRibbonButton();
            this.btnOpOrder = this.Factory.CreateRibbonButton();
            this.btnWarningOrder = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.btnCaseStudy = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.AppxC.SuspendLayout();
            this.group2.SuspendLayout();
            this.group1.SuspendLayout();
            this.group3.SuspendLayout();
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
            this.AppxC.Groups.Add(this.group1);
            this.AppxC.Groups.Add(this.group3);
            this.AppxC.Label = "JSSD";
            this.AppxC.Name = "AppxC";
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnEnable);
            this.group2.Items.Add(this.btnDisable);
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
            // group1
            // 
            this.group1.Items.Add(this.btnReplaceAll);
            this.group1.Items.Add(this.btnHighlightAll);
            this.group1.Name = "group1";
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
            // group3
            // 
            this.group3.Items.Add(this.menuTemplates);
            this.group3.Items.Add(this.button1);
            this.group3.Name = "group3";
            // 
            // menuTemplates
            // 
            this.menuTemplates.Items.Add(this.btnGeneralFormat);
            this.menuTemplates.Items.Add(this.btnDoLetter);
            this.menuTemplates.Items.Add(this.btnNotingSheet);
            this.menuTemplates.Items.Add(this.btnSignalForm);
            this.menuTemplates.Items.Add(this.btnStatementOfCase);
            this.menuTemplates.Items.Add(this.btnStatementOfCaseDPM);
            this.menuTemplates.Items.Add(this.btnAppxFormat);
            this.menuTemplates.Items.Add(this.btnServicePaper);
            this.menuTemplates.Items.Add(this.btnAgendaPts);
            this.menuTemplates.Items.Add(this.btnOpNotes);
            this.menuTemplates.Items.Add(this.btnMoM);
            this.menuTemplates.Items.Add(this.btnEmailFormat);
            this.menuTemplates.Items.Add(this.btnTourNotes);
            this.menuTemplates.Items.Add(this.btnAppreciation);
            this.menuTemplates.Items.Add(this.btnOpOrder);
            this.menuTemplates.Items.Add(this.btnWarningOrder);
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
            // btnDoLetter
            // 
            this.btnDoLetter.Label = "DO Letter";
            this.btnDoLetter.Name = "btnDoLetter";
            this.btnDoLetter.ShowImage = true;
            this.btnDoLetter.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDoLetter_Click);
            // 
            // btnNotingSheet
            // 
            this.btnNotingSheet.Label = "Noting Sheet";
            this.btnNotingSheet.Name = "btnNotingSheet";
            this.btnNotingSheet.ShowImage = true;
            this.btnNotingSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNotingSheet_Click);
            // 
            // btnSignalForm
            // 
            this.btnSignalForm.Label = "Signal Form";
            this.btnSignalForm.Name = "btnSignalForm";
            this.btnSignalForm.ShowImage = true;
            this.btnSignalForm.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSignalForm_Click);
            // 
            // btnStatementOfCase
            // 
            this.btnStatementOfCase.Label = "Statement of Case";
            this.btnStatementOfCase.Name = "btnStatementOfCase";
            this.btnStatementOfCase.ShowImage = true;
            this.btnStatementOfCase.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStatementOfCase_Click);
            // 
            // btnStatementOfCaseDPM
            // 
            this.btnStatementOfCaseDPM.Label = "Statement of Case (DPM)";
            this.btnStatementOfCaseDPM.Name = "btnStatementOfCaseDPM";
            this.btnStatementOfCaseDPM.ShowImage = true;
            this.btnStatementOfCaseDPM.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStatementOfCaseDPM_Click);
            // 
            // btnAppxFormat
            // 
            this.btnAppxFormat.Label = "Appx Format";
            this.btnAppxFormat.Name = "btnAppxFormat";
            this.btnAppxFormat.ShowImage = true;
            this.btnAppxFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAppxFormat_Click);
            // 
            // btnServicePaper
            // 
            this.btnServicePaper.Label = "Service Paper";
            this.btnServicePaper.Name = "btnServicePaper";
            this.btnServicePaper.ShowImage = true;
            this.btnServicePaper.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnServicePaper_Click);
            // 
            // btnAgendaPts
            // 
            this.btnAgendaPts.Label = "Agenda Pts";
            this.btnAgendaPts.Name = "btnAgendaPts";
            this.btnAgendaPts.ShowImage = true;
            this.btnAgendaPts.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAgendaPts_Click);
            // 
            // btnOpNotes
            // 
            this.btnOpNotes.Label = "Op Notes";
            this.btnOpNotes.Name = "btnOpNotes";
            this.btnOpNotes.ShowImage = true;
            this.btnOpNotes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOpNotes_Click);
            // 
            // btnMoM
            // 
            this.btnMoM.Label = "MoM";
            this.btnMoM.Name = "btnMoM";
            this.btnMoM.ShowImage = true;
            this.btnMoM.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMoM_Click);
            // 
            // btnEmailFormat
            // 
            this.btnEmailFormat.Label = "E-mail Format";
            this.btnEmailFormat.Name = "btnEmailFormat";
            this.btnEmailFormat.ShowImage = true;
            this.btnEmailFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEmailFormat_Click);
            // 
            // btnTourNotes
            // 
            this.btnTourNotes.Label = "Tour Notes";
            this.btnTourNotes.Name = "btnTourNotes";
            this.btnTourNotes.ShowImage = true;
            this.btnTourNotes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTourNotes_Click);
            // 
            // btnAppreciation
            // 
            this.btnAppreciation.Label = "Appreciation";
            this.btnAppreciation.Name = "btnAppreciation";
            this.btnAppreciation.ShowImage = true;
            this.btnAppreciation.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAppreciation_Click);
            // 
            // btnOpOrder
            // 
            this.btnOpOrder.Label = "Op Order";
            this.btnOpOrder.Name = "btnOpOrder";
            this.btnOpOrder.ShowImage = true;
            this.btnOpOrder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOpOrder_Click);
            // 
            // btnWarningOrder
            // 
            this.btnWarningOrder.Label = "Warning Order";
            this.btnWarningOrder.Name = "btnWarningOrder";
            this.btnWarningOrder.ShowImage = true;
            this.btnWarningOrder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnWarningOrder_Click);
            // 
            // button1
            // 
            this.button1.Label = "Show Suggestions";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShowSuggestions_Click);
            // 
            // btnCaseStudy
            // 
            this.btnCaseStudy.Label = "Case Study";
            this.btnCaseStudy.Name = "btnCaseStudy";
            this.btnCaseStudy.ShowImage = true;
            this.btnCaseStudy.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCaseStudy_Click);
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
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }


        private void btnGeneralFormat_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.General Format.docx", "General Format.docx");
            InsertTemplate(templatePath);
        }


        private void btnDoLetter_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.DO Letter.docx", "DO Letter.docx");
            InsertTemplate(templatePath);
        }

        private void btnNotingSheet_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Noting Sheet.docx", "Noting Sheet.docx");
            InsertTemplate(templatePath);
        }

        private void btnSignalForm_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Signal Form.docx", "Signal Form.docx");
            InsertTemplate(templatePath);
        }

        private void btnStatementOfCase_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Statement of Case.docx", "Statement of Case.docx");
            InsertTemplate(templatePath);
        }

        private void btnStatementOfCaseDPM_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Statement of Case (DPM).docx", "Statement of Case (DPM).docx");
            InsertTemplate(templatePath);
        }

        private void btnAppxFormat_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Appx Format.docx", "Appx Format.docx");
            InsertTemplate(templatePath);
        }

        private void btnServicePaper_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Service Paper.docx", "Service Paper.docx");
            InsertTemplate(templatePath);
        }

        private void btnAgendaPts_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Agenda Pts.docx", "Agenda Pts.docx");
            InsertTemplate(templatePath);
        }

        private void btnOpNotes_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Op Notes.docx", "Op Notes.docx");
            InsertTemplate(templatePath);
        }

        private void btnMoM_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.MoM.docx", "MoM.docx");
            InsertTemplate(templatePath);
        }

        private void btnEmailFormat_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.E-mail Format.docx", "E-mail Format.docx");
            InsertTemplate(templatePath);
        }

        private void btnTourNotes_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Tour Notes.docx", "Tour Notes.docx");
            InsertTemplate(templatePath);
        }

        private void btnAppreciation_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Appreciation.docx", "Appreciation.docx");
            InsertTemplate(templatePath);
        }

        private void btnCaseStudy_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Case Study.docx", "Case Study.docx");
            InsertTemplate(templatePath);
        }

        private void btnOpOrder_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Op Order.docx", "Op Order.docx");
            InsertTemplate(templatePath);
        }

        private void btnWarningOrder_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Warning Order.docx", "Warning Order.docx");
            InsertTemplate(templatePath);
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

       

        private void InsertTemplate(string templatePath)
        {
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDoLetter;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNotingSheet;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSignalForm;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStatementOfCase;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStatementOfCaseDPM;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAppxFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnServicePaper;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAgendaPts;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOpNotes;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMoM;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEmailFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTourNotes;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAppreciation;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCaseStudy;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOpOrder;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnWarningOrder;
        internal RibbonButton button1;
        internal RibbonGroup group1;
        internal RibbonGroup group3;
    }

    partial class ThisRibbonCollection
    {
        internal AbbreviationRibbon AbbreviationRibbon
        {
            get { return this.GetRibbon<AbbreviationRibbon>(); }
        }
    }
}
