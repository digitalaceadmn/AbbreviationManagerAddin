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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AbbreviationRibbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.JSSD = this.Factory.CreateRibbonTab();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnEnable = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.btnDisable = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnReplaceAll = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.btnHighlightAll = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.menuTemplates = this.Factory.CreateRibbonMenu();
            this.GovtService = this.Factory.CreateRibbonMenu();
            this.btnGovtofindialetters = this.Factory.CreateRibbonButton();
            this.btnGovtofindia = this.Factory.CreateRibbonButton();
            this.btnServiceletters = this.Factory.CreateRibbonButton();
            this.btnEmailFormat = this.Factory.CreateRibbonButton();
            this.btnSignalFormat = this.Factory.CreateRibbonButton();
            this.btnNoteSheet = this.Factory.CreateRibbonButton();
            this.btnDoLetter = this.Factory.CreateRibbonButton();
            this.StaffPaper = this.Factory.CreateRibbonMenu();
            this.btnAppreciation = this.Factory.CreateRibbonButton();
            this.btnCopp = this.Factory.CreateRibbonButton();
            this.btnServicePaper = this.Factory.CreateRibbonButton();
            this.btnProblemStatement = this.Factory.CreateRibbonButton();
            this.btnStatementOfCase = this.Factory.CreateRibbonButton();
            this.btnStatementOfCaseDPM = this.Factory.CreateRibbonButton();
            this.menu1 = this.Factory.CreateRibbonMenu();
            this.btnAgendaPts = this.Factory.CreateRibbonButton();
            this.btnMoMeeting = this.Factory.CreateRibbonButton();
            this.btnBrief = this.Factory.CreateRibbonButton();
            this.btnReturnBrief = this.Factory.CreateRibbonButton();
            this.btnTourNotes = this.Factory.CreateRibbonButton();
            this.btnNotice = this.Factory.CreateRibbonButton();
            this.btnCabinetNote = this.Factory.CreateRibbonButton();
            this.btnPressRelease = this.Factory.CreateRibbonButton();
            this.btnSocialMediaPost = this.Factory.CreateRibbonButton();
            this.btnGazetteNotification = this.Factory.CreateRibbonButton();
            this.menu2 = this.Factory.CreateRibbonMenu();
            this.btnParliamentaryQuestionForwarding = this.Factory.CreateRibbonButton();
            this.btnParliamentaryQuestionReply = this.Factory.CreateRibbonButton();
            this.btnCoveringletterReply = this.Factory.CreateRibbonButton();
            this.menu3 = this.Factory.CreateRibbonMenu();
            this.btnSampleWarningOrder = this.Factory.CreateRibbonButton();
            this.btnOpOrder = this.Factory.CreateRibbonButton();
            this.menu4 = this.Factory.CreateRibbonMenu();
            this.btnMOMemorandum = this.Factory.CreateRibbonButton();
            this.btnBookReview = this.Factory.CreateRibbonButton();
            this.btnAnnotationinCorrespondence = this.Factory.CreateRibbonButton();
            this.btnWarningOrder = this.Factory.CreateRibbonButton();
            this.separator4 = this.Factory.CreateRibbonSeparator();
            this.button1 = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.button2 = this.Factory.CreateRibbonButton();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.highLightLike = this.Factory.CreateRibbonButton();
            this.btnCaseStudy = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.JSSD.SuspendLayout();
            this.group2.SuspendLayout();
            this.group1.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // JSSD
            // 
            this.JSSD.Groups.Add(this.group2);
            this.JSSD.Groups.Add(this.group1);
            this.JSSD.Groups.Add(this.group3);
            this.JSSD.Groups.Add(this.group4);
            this.JSSD.Label = "JSSD";
            this.JSSD.Name = "JSSD";
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnEnable);
            this.group2.Items.Add(this.separator1);
            this.group2.Items.Add(this.btnDisable);
            this.group2.Label = "Enable / Disable";
            this.group2.Name = "group2";
            this.group2.Visible = false;

            this.JSSD_New = this.Factory.CreateRibbonTab();
            this.JSSD_New.Groups.Add(this.group2);
            this.JSSD_New.Groups.Add(this.group1);
            this.JSSD_New.Groups.Add(this.group3);
            this.JSSD_New.Groups.Add(this.group4);
            this.JSSD_New.Label = "JSSD (New)";
            this.JSSD_New.Name = "JSSD_New";
            // 
            // btnEnable
            // 
            this.btnEnable.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnEnable.Image = ((System.Drawing.Image)(resources.GetObject("btnEnable.Image")));
            this.btnEnable.Label = "Enable Abbreviation";
            this.btnEnable.Name = "btnEnable";
            this.btnEnable.ShowImage = true;
            this.btnEnable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEnable_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // btnDisable
            // 
            this.btnDisable.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDisable.Image = ((System.Drawing.Image)(resources.GetObject("btnDisable.Image")));
            this.btnDisable.Label = "Disable Abbreviation";
            this.btnDisable.Name = "btnDisable";
            this.btnDisable.ShowImage = true;
            this.btnDisable.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDisable_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnReplaceAll);
            this.group1.Items.Add(this.separator2);
            this.group1.Items.Add(this.btnHighlightAll);
            this.group1.Label = "Replace / Highlight";
            this.group1.Name = "group1";
            // 
            // btnReplaceAll
            // 
            this.btnReplaceAll.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnReplaceAll.Image = ((System.Drawing.Image)(resources.GetObject("btnReplaceAll.Image")));
            this.btnReplaceAll.Label = "Replace All";
            this.btnReplaceAll.Name = "btnReplaceAll";
            this.btnReplaceAll.ShowImage = true;
            this.btnReplaceAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReplaceAll_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // btnHighlightAll
            // 
            this.btnHighlightAll.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnHighlightAll.Image = ((System.Drawing.Image)(resources.GetObject("btnHighlightAll.Image")));
            this.btnHighlightAll.Label = "Highlight All";
            this.btnHighlightAll.Name = "btnHighlightAll";
            this.btnHighlightAll.ShowImage = true;
            this.btnHighlightAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnHighlightAll_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.menuTemplates);
            this.group3.Items.Add(this.separator4);
            this.group3.Items.Add(this.button1);
            this.group3.Label = "Template / Show Suggestions";
            this.group3.Name = "group3";
            // 
            // menuTemplates
            // 
            this.menuTemplates.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menuTemplates.Image = ((System.Drawing.Image)(resources.GetObject("menuTemplates.Image")));
            this.menuTemplates.Items.Add(this.GovtService);
            this.menuTemplates.Items.Add(this.StaffPaper);
            this.menuTemplates.Items.Add(this.menu1);
            this.menuTemplates.Items.Add(this.menu2);
            this.menuTemplates.Items.Add(this.menu3);
            this.menuTemplates.Items.Add(this.menu4);
            this.menuTemplates.Items.Add(this.btnWarningOrder);
            this.menuTemplates.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menuTemplates.Label = "Template";
            this.menuTemplates.Name = "menuTemplates";
            this.menuTemplates.ShowImage = true;
            // 
            // GovtService
            // 
            this.GovtService.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.GovtService.Image = ((System.Drawing.Image)(resources.GetObject("GovtService.Image")));
            this.GovtService.Items.Add(this.btnGovtofindialetters);
            this.GovtService.Items.Add(this.btnGovtofindia);
            this.GovtService.Items.Add(this.btnServiceletters);
            this.GovtService.Items.Add(this.btnEmailFormat);
            this.GovtService.Items.Add(this.btnSignalFormat);
            this.GovtService.Items.Add(this.btnNoteSheet);
            this.GovtService.Items.Add(this.btnDoLetter);
            this.GovtService.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.GovtService.Label = "Govt & Service Correspondence";
            this.GovtService.Name = "GovtService";
            this.GovtService.ShowImage = true;
            // 
            // btnGovtofindialetters
            // 
            this.btnGovtofindialetters.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnGovtofindialetters.Image = ((System.Drawing.Image)(resources.GetObject("btnGovtofindialetters.Image")));
            this.btnGovtofindialetters.Label = "Govt of India letters";
            this.btnGovtofindialetters.Name = "btnGovtofindialetters";
            this.btnGovtofindialetters.ShowImage = true;
            this.btnGovtofindialetters.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGovtofindialetters_Click);
            // 
            // btnGovtofindia
            // 
            this.btnGovtofindia.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnGovtofindia.Image = ((System.Drawing.Image)(resources.GetObject("btnGovtofindia.Image")));
            this.btnGovtofindia.Label = "Govt of India (Inter Departmental Note)";
            this.btnGovtofindia.Name = "btnGovtofindia";
            this.btnGovtofindia.ShowImage = true;
            this.btnGovtofindia.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGovtofindia_Click);
            // 
            // btnServiceletters
            // 
            this.btnServiceletters.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnServiceletters.Image = ((System.Drawing.Image)(resources.GetObject("btnServiceletters.Image")));
            this.btnServiceletters.Label = "Service letters";
            this.btnServiceletters.Name = "btnServiceletters";
            this.btnServiceletters.ShowImage = true;
            this.btnServiceletters.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnServiceletters_Click);
            // 
            // btnEmailFormat
            // 
            this.btnEmailFormat.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnEmailFormat.Image = ((System.Drawing.Image)(resources.GetObject("btnEmailFormat.Image")));
            this.btnEmailFormat.Label = "E-mail Format";
            this.btnEmailFormat.Name = "btnEmailFormat";
            this.btnEmailFormat.ShowImage = true;
            this.btnEmailFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEmailFormat_Click);
            // 
            // btnSignalFormat
            // 
            this.btnSignalFormat.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSignalFormat.Image = ((System.Drawing.Image)(resources.GetObject("btnSignalFormat.Image")));
            this.btnSignalFormat.Label = "Signal Format";
            this.btnSignalFormat.Name = "btnSignalFormat";
            this.btnSignalFormat.ShowImage = true;
            this.btnSignalFormat.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSignalForm_Click);
            // 
            // btnNoteSheet
            // 
            this.btnNoteSheet.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnNoteSheet.Image = ((System.Drawing.Image)(resources.GetObject("btnNoteSheet.Image")));
            this.btnNoteSheet.Label = "Note Sheet";
            this.btnNoteSheet.Name = "btnNoteSheet";
            this.btnNoteSheet.ShowImage = true;
            this.btnNoteSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNoteSheet_Click);
            // 
            // btnDoLetter
            // 
            this.btnDoLetter.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDoLetter.Image = ((System.Drawing.Image)(resources.GetObject("btnDoLetter.Image")));
            this.btnDoLetter.Label = "DO Letter";
            this.btnDoLetter.Name = "btnDoLetter";
            this.btnDoLetter.ShowImage = true;
            this.btnDoLetter.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDoLetter_Click);
            // 
            // StaffPaper
            // 
            this.StaffPaper.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.StaffPaper.Image = ((System.Drawing.Image)(resources.GetObject("StaffPaper.Image")));
            this.StaffPaper.Items.Add(this.btnAppreciation);
            this.StaffPaper.Items.Add(this.btnCopp);
            this.StaffPaper.Items.Add(this.btnServicePaper);
            this.StaffPaper.Items.Add(this.btnProblemStatement);
            this.StaffPaper.Items.Add(this.btnStatementOfCase);
            this.StaffPaper.Items.Add(this.btnStatementOfCaseDPM);
            this.StaffPaper.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.StaffPaper.Label = "Staff Paper";
            this.StaffPaper.Name = "StaffPaper";
            this.StaffPaper.ShowImage = true;
            // 
            // btnAppreciation
            // 
            this.btnAppreciation.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAppreciation.Image = ((System.Drawing.Image)(resources.GetObject("btnAppreciation.Image")));
            this.btnAppreciation.Label = "Appreciation";
            this.btnAppreciation.Name = "btnAppreciation";
            this.btnAppreciation.ShowImage = true;
            this.btnAppreciation.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAppreciation_Click);
            // 
            // btnCopp
            // 
            this.btnCopp.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCopp.Image = ((System.Drawing.Image)(resources.GetObject("btnCopp.Image")));
            this.btnCopp.Label = "COPP";
            this.btnCopp.Name = "btnCopp";
            this.btnCopp.ShowImage = true;
            this.btnCopp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCopp_Click);
            // 
            // btnServicePaper
            // 
            this.btnServicePaper.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnServicePaper.Image = ((System.Drawing.Image)(resources.GetObject("btnServicePaper.Image")));
            this.btnServicePaper.Label = "Service Paper";
            this.btnServicePaper.Name = "btnServicePaper";
            this.btnServicePaper.ShowImage = true;
            this.btnServicePaper.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnServicePaper_Click);
            // 
            // btnProblemStatement
            // 
            this.btnProblemStatement.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnProblemStatement.Image = ((System.Drawing.Image)(resources.GetObject("btnProblemStatement.Image")));
            this.btnProblemStatement.Label = "Problem Statement";
            this.btnProblemStatement.Name = "btnProblemStatement";
            this.btnProblemStatement.ShowImage = true;
            this.btnProblemStatement.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnProblemStatement_Click);
            // 
            // btnStatementOfCase
            // 
            this.btnStatementOfCase.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnStatementOfCase.Image = ((System.Drawing.Image)(resources.GetObject("btnStatementOfCase.Image")));
            this.btnStatementOfCase.Label = "Statement of Case";
            this.btnStatementOfCase.Name = "btnStatementOfCase";
            this.btnStatementOfCase.ShowImage = true;
            this.btnStatementOfCase.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStatementOfCase_Click);
            // 
            // btnStatementOfCaseDPM
            // 
            this.btnStatementOfCaseDPM.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnStatementOfCaseDPM.Image = ((System.Drawing.Image)(resources.GetObject("btnStatementOfCaseDPM.Image")));
            this.btnStatementOfCaseDPM.Label = "Statement of Case (DPM)";
            this.btnStatementOfCaseDPM.Name = "btnStatementOfCaseDPM";
            this.btnStatementOfCaseDPM.ShowImage = true;
            this.btnStatementOfCaseDPM.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStatementOfCaseDPM_Click);
            // 
            // menu1
            // 
            this.menu1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menu1.Image = ((System.Drawing.Image)(resources.GetObject("menu1.Image")));
            this.menu1.Items.Add(this.btnAgendaPts);
            this.menu1.Items.Add(this.btnMoMeeting);
            this.menu1.Items.Add(this.btnBrief);
            this.menu1.Items.Add(this.btnReturnBrief);
            this.menu1.Items.Add(this.btnTourNotes);
            this.menu1.Items.Add(this.btnNotice);
            this.menu1.Items.Add(this.btnCabinetNote);
            this.menu1.Items.Add(this.btnPressRelease);
            this.menu1.Items.Add(this.btnSocialMediaPost);
            this.menu1.Items.Add(this.btnGazetteNotification);
            this.menu1.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menu1.Label = "Note / Notices";
            this.menu1.Name = "menu1";
            this.menu1.ShowImage = true;
            // 
            // btnAgendaPts
            // 
            this.btnAgendaPts.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAgendaPts.Image = ((System.Drawing.Image)(resources.GetObject("btnAgendaPts.Image")));
            this.btnAgendaPts.Label = "Agenda Pts";
            this.btnAgendaPts.Name = "btnAgendaPts";
            this.btnAgendaPts.ShowImage = true;
            this.btnAgendaPts.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAgendaPts_Click);
            // 
            // btnMoMeeting
            // 
            this.btnMoMeeting.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnMoMeeting.Image = ((System.Drawing.Image)(resources.GetObject("btnMoMeeting.Image")));
            this.btnMoMeeting.Label = "MOM (Minutes of Meeting)";
            this.btnMoMeeting.Name = "btnMoMeeting";
            this.btnMoMeeting.ShowImage = true;
            this.btnMoMeeting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMoMeeting_Click);
            // 
            // btnBrief
            // 
            this.btnBrief.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnBrief.Image = ((System.Drawing.Image)(resources.GetObject("btnBrief.Image")));
            this.btnBrief.Label = "Brief";
            this.btnBrief.Name = "btnBrief";
            this.btnBrief.ShowImage = true;
            this.btnBrief.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnBrief_Click);
            // 
            // btnReturnBrief
            // 
            this.btnReturnBrief.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnReturnBrief.Image = ((System.Drawing.Image)(resources.GetObject("btnReturnBrief.Image")));
            this.btnReturnBrief.Label = "Return Brief";
            this.btnReturnBrief.Name = "btnReturnBrief";
            this.btnReturnBrief.ShowImage = true;
            this.btnReturnBrief.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReturnBrief_Click);
            // 
            // btnTourNotes
            // 
            this.btnTourNotes.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnTourNotes.Image = ((System.Drawing.Image)(resources.GetObject("btnTourNotes.Image")));
            this.btnTourNotes.Label = "Tour Notes";
            this.btnTourNotes.Name = "btnTourNotes";
            this.btnTourNotes.ShowImage = true;
            this.btnTourNotes.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTourNotes_Click);
            // 
            // btnNotice
            // 
            this.btnNotice.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnNotice.Image = ((System.Drawing.Image)(resources.GetObject("btnNotice.Image")));
            this.btnNotice.Label = "Notice";
            this.btnNotice.Name = "btnNotice";
            this.btnNotice.ShowImage = true;
            this.btnNotice.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNotice_Click);
            // 
            // btnCabinetNote
            // 
            this.btnCabinetNote.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCabinetNote.Image = ((System.Drawing.Image)(resources.GetObject("btnCabinetNote.Image")));
            this.btnCabinetNote.Label = "Cabinet Note";
            this.btnCabinetNote.Name = "btnCabinetNote";
            this.btnCabinetNote.ShowImage = true;
            this.btnCabinetNote.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCabinetNote_Click);
            // 
            // btnPressRelease
            // 
            this.btnPressRelease.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnPressRelease.Image = ((System.Drawing.Image)(resources.GetObject("btnPressRelease.Image")));
            this.btnPressRelease.Label = "Press Release";
            this.btnPressRelease.Name = "btnPressRelease";
            this.btnPressRelease.ShowImage = true;
            this.btnPressRelease.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPressRelease_Click);
            // 
            // btnSocialMediaPost
            // 
            this.btnSocialMediaPost.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSocialMediaPost.Image = ((System.Drawing.Image)(resources.GetObject("btnSocialMediaPost.Image")));
            this.btnSocialMediaPost.Label = "Social Media Post";
            this.btnSocialMediaPost.Name = "btnSocialMediaPost";
            this.btnSocialMediaPost.ShowImage = true;
            this.btnSocialMediaPost.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSocialMediaPost_Click);
            // 
            // btnGazetteNotification
            // 
            this.btnGazetteNotification.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnGazetteNotification.Image = ((System.Drawing.Image)(resources.GetObject("btnGazetteNotification.Image")));
            this.btnGazetteNotification.Label = "Gazette Notification";
            this.btnGazetteNotification.Name = "btnGazetteNotification";
            this.btnGazetteNotification.ShowImage = true;
            this.btnGazetteNotification.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGazetteNotification_Click);
            // 
            // menu2
            // 
            this.menu2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menu2.Image = ((System.Drawing.Image)(resources.GetObject("menu2.Image")));
            this.menu2.Items.Add(this.btnParliamentaryQuestionForwarding);
            this.menu2.Items.Add(this.btnParliamentaryQuestionReply);
            this.menu2.Items.Add(this.btnCoveringletterReply);
            this.menu2.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menu2.Label = "Parliamentary Question";
            this.menu2.Name = "menu2";
            this.menu2.ShowImage = true;
            // 
            // btnParliamentaryQuestionForwarding
            // 
            this.btnParliamentaryQuestionForwarding.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnParliamentaryQuestionForwarding.Image = ((System.Drawing.Image)(resources.GetObject("btnParliamentaryQuestionForwarding.Image")));
            this.btnParliamentaryQuestionForwarding.Label = "Parliamentary Question (Forwarding)";
            this.btnParliamentaryQuestionForwarding.Name = "btnParliamentaryQuestionForwarding";
            this.btnParliamentaryQuestionForwarding.ShowImage = true;
            this.btnParliamentaryQuestionForwarding.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnParliamentaryQuestionForwarding_Click);
            // 
            // btnParliamentaryQuestionReply
            // 
            this.btnParliamentaryQuestionReply.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnParliamentaryQuestionReply.Image = ((System.Drawing.Image)(resources.GetObject("btnParliamentaryQuestionReply.Image")));
            this.btnParliamentaryQuestionReply.Label = "Parliamentary Question (Reply)";
            this.btnParliamentaryQuestionReply.Name = "btnParliamentaryQuestionReply";
            this.btnParliamentaryQuestionReply.ShowImage = true;
            this.btnParliamentaryQuestionReply.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnParliamentaryQuestionReply_Click);
            // 
            // btnCoveringletterReply
            // 
            this.btnCoveringletterReply.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCoveringletterReply.Image = ((System.Drawing.Image)(resources.GetObject("btnCoveringletterReply.Image")));
            this.btnCoveringletterReply.Label = "Covering letter (Reply)";
            this.btnCoveringletterReply.Name = "btnCoveringletterReply";
            this.btnCoveringletterReply.ShowImage = true;
            this.btnCoveringletterReply.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCoveringletterReply_Click);
            // 
            // menu3
            // 
            this.menu3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menu3.Image = ((System.Drawing.Image)(resources.GetObject("menu3.Image")));
            this.menu3.Items.Add(this.btnSampleWarningOrder);
            this.menu3.Items.Add(this.btnOpOrder);
            this.menu3.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menu3.Label = "Directives / Order / Instr";
            this.menu3.Name = "menu3";
            this.menu3.ShowImage = true;
            // 
            // btnSampleWarningOrder
            // 
            this.btnSampleWarningOrder.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSampleWarningOrder.Image = ((System.Drawing.Image)(resources.GetObject("btnSampleWarningOrder.Image")));
            this.btnSampleWarningOrder.Label = "Sample Warning Order";
            this.btnSampleWarningOrder.Name = "btnSampleWarningOrder";
            this.btnSampleWarningOrder.ShowImage = true;
            this.btnSampleWarningOrder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSampleWarningOrder_Click);
            // 
            // btnOpOrder
            // 
            this.btnOpOrder.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnOpOrder.Image = ((System.Drawing.Image)(resources.GetObject("btnOpOrder.Image")));
            this.btnOpOrder.Label = "Op Order";
            this.btnOpOrder.Name = "btnOpOrder";
            this.btnOpOrder.ShowImage = true;
            this.btnOpOrder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOpOrder_Click);
            // 
            // menu4
            // 
            this.menu4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menu4.Image = ((System.Drawing.Image)(resources.GetObject("menu4.Image")));
            this.menu4.Items.Add(this.btnMOMemorandum);
            this.menu4.Items.Add(this.btnBookReview);
            this.menu4.Items.Add(this.btnAnnotationinCorrespondence);
            this.menu4.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menu4.Label = "Misc";
            this.menu4.Name = "menu4";
            this.menu4.ShowImage = true;
            // 
            // btnMOMemorandum
            // 
            this.btnMOMemorandum.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnMOMemorandum.Image = ((System.Drawing.Image)(resources.GetObject("btnMOMemorandum.Image")));
            this.btnMOMemorandum.Label = "MOM (Memorandum of Understanding)";
            this.btnMOMemorandum.Name = "btnMOMemorandum";
            this.btnMOMemorandum.ShowImage = true;
            this.btnMOMemorandum.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMOMemorandum_Click);
            // 
            // btnBookReview
            // 
            this.btnBookReview.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnBookReview.Image = ((System.Drawing.Image)(resources.GetObject("btnBookReview.Image")));
            this.btnBookReview.Label = "Book Review";
            this.btnBookReview.Name = "btnBookReview";
            this.btnBookReview.ShowImage = true;
            this.btnBookReview.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnBookReview_Click);
            // 
            // btnAnnotationinCorrespondence
            // 
            this.btnAnnotationinCorrespondence.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAnnotationinCorrespondence.Image = ((System.Drawing.Image)(resources.GetObject("btnAnnotationinCorrespondence.Image")));
            this.btnAnnotationinCorrespondence.Label = "Annotation in Correspondence";
            this.btnAnnotationinCorrespondence.Name = "btnAnnotationinCorrespondence";
            this.btnAnnotationinCorrespondence.ShowImage = true;
            this.btnAnnotationinCorrespondence.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAnnotationinCorrespondence_Click);
            // 
            // btnWarningOrder
            // 
            this.btnWarningOrder.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnWarningOrder.Image = ((System.Drawing.Image)(resources.GetObject("btnWarningOrder.Image")));
            this.btnWarningOrder.Label = "Key Features (JSSD)";
            this.btnWarningOrder.Name = "btnWarningOrder";
            this.btnWarningOrder.ShowImage = true;
            this.btnWarningOrder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnWarningOrder_Click);
            // 
            // separator4
            // 
            this.separator4.Name = "separator4";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Label = "Show Suggestions";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ShowSuggestions_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.button2);
            this.group4.Label = "Help";
            this.group4.Name = "group4";
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
            this.button2.Label = "Help";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // highLightLike
            // 
            this.highLightLike.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.highLightLike.Image = ((System.Drawing.Image)(resources.GetObject("highLightLike.Image")));
            this.highLightLike.Label = "Highlight Like";
            this.highLightLike.Name = "highLightLike";
            this.highLightLike.ShowImage = true;
            this.highLightLike.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.highLightLike_Click);
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
            this.Tabs.Add(this.JSSD);
            this.Tabs.Add(this.JSSD_New);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.AbbreviationRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.JSSD.ResumeLayout(false);
            this.JSSD.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.ResumeLayout(false);

        }


        private void btnGovtofindialetters_Click(object sender, RibbonControlEventArgs e)
        {
            
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Govt of India letters.docx", "Govt of India letters.docx");
            InsertTemplate(templatePath);
        }

        private void btnGovtofindia_Click(object sender, RibbonControlEventArgs e)
        {
            
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Govt of India (Inter Departmental Note).docx", "Govt of India (Inter Departmental Note).docx");
            InsertTemplate(templatePath);
        }

        private void btnServiceletters_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Service letters.docx", "Service letters.docx");
            InsertTemplate(templatePath);
        }

        private void btnEmailFormat_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.E-mail Format.docx", "E-mail Format.docx");
            InsertTemplate(templatePath);
        }

        private void btnSignalForm_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Signal Form.docx", "Signal Form.docx");
            InsertTemplate(templatePath);
        }

        private void btnNoteSheet_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Note Sheet.docx", "Note Sheet.docx");
            InsertTemplate(templatePath);
        }


        private void btnDoLetter_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.DO Letter.docx", "DO Letter.docx");
            InsertTemplate(templatePath);
        }

        private void btnAppreciation_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Appreciation.docx", "Appreciation.docx");
            InsertTemplate(templatePath);
        }


        private void btnCopp_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.COPP.docx", "COPP.docx");
            InsertTemplate(templatePath);
        }

        private void btnServicePaper_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Service Paper.docx", "Service Paper.docx");
            InsertTemplate(templatePath);
        }

        private void btnProblemStatement_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Problem Statement.docx", "Problem Statement.docx");
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

        private void btnAgendaPts_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Agenda Pts.docx", "Agenda Pts.docx");
            InsertTemplate(templatePath);
        }


        private void btnMoMeeting_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.MOM (Minutes of Meeting).docx", "MOM (Minutes of Meeting).docx");
            InsertTemplate(templatePath);
        }

        private void btnBrief_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Brief.docx", "Brief.docx");
            InsertTemplate(templatePath);
        }

        private void btnReturnBrief_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Return Brief.docx", "Return Brief.docx");
            InsertTemplate(templatePath);
        }

        private void btnTourNotes_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Tour Notes.docx", "Tour Notes.docx");
            InsertTemplate(templatePath);
        }


        private void btnNotice_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Notice.docx", "Notice.docx");
            InsertTemplate(templatePath);
        }

        private void btnCabinetNote_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Cabinet Note.docx", "Cabinet Note.docx");
            InsertTemplate(templatePath);
        }

        private void btnPressRelease_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Press Release.docx", "Press Release.docx");
            InsertTemplate(templatePath);
        }

        private void btnSocialMediaPost_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Social Media Post.docx", "Social Media Post.docx");
            InsertTemplate(templatePath);
        }

        private void btnGazetteNotification_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Gazette Notification.docx", "Gazette Notification.docx");
            InsertTemplate(templatePath);
        }

        private void btnParliamentaryQuestionForwarding_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Parliamentary Question (Forwarding).docx", "Parliamentary Question (Forwarding).docx");
            InsertTemplate(templatePath);
        }


        private void btnParliamentaryQuestionReply_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Parliamentary Question (Reply).docx", "Parliamentary Question (Reply).docx");
            InsertTemplate(templatePath);
        }


        private void btnCoveringletterReply_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Covering letter (Reply).docx", "Covering letter (Reply).docx");
            InsertTemplate(templatePath);
        }

        private void btnSampleWarningOrder_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Simple Warning Order.docx", "Simple Warning Order.docx");
            InsertTemplate(templatePath);
        }

        private void btnOpOrder_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Op order.docx", "Op order.docx");
            InsertTemplate(templatePath);
        }

        private void btnMOMemorandum_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.MOM (Memorandum of Understanding).docx", "MOM (Memorandum of Understanding).docx");
            InsertTemplate(templatePath);
        }

        private void btnBookReview_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Book Review.docx", "Book Review.docx");
            InsertTemplate(templatePath);
        }


        private void btnAnnotationinCorrespondence_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Annotation in Correspondence.docx", "Annotation in Correspondence.docx");
            InsertTemplate(templatePath);
        }


        private void btnGeneralFormat_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.General Format.docx", "General Format.docx");
            InsertTemplate(templatePath);
        }

        private void btnNotingSheet_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Noting Sheet.docx", "Noting Sheet.docx");
            InsertTemplate(templatePath);
        }

        

        private void btnAppxFormat_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Appx Format.docx", "Appx Format.docx");
            InsertTemplate(templatePath);
        }

        

        

        private void btnOpNotes_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Op Notes.docx", "Op Notes.docx");
            InsertTemplate(templatePath);
        }

        

       

        

        

        

        

        private void btnCaseStudy_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Case Study.docx", "Case Study.docx");
            InsertTemplate(templatePath);
        }

        

        private void btnWarningOrder_Click(object sender, RibbonControlEventArgs e)
        {
            string templatePath = ExtractTemplateToLocal("AbbreviationWordAddin.Templates.Key Features (JSSD).docx", "Key Features (JSSD).docx");
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

                tempDoc.Content.Copy();

                tempDoc.Close(false);

                System.Threading.Thread.Sleep(200);

                var currentDoc = Globals.ThisAddIn.Application.ActiveDocument;
                currentDoc.Activate();

                var selection = Globals.ThisAddIn.Application.Selection;

                selection.HomeKey(WdUnits.wdStory);

                selection.Paste();

                selection.TypeParagraph();

                selection.HomeKey(WdUnits.wdStory);

                MessageBox.Show("Template inserted successfully at the top!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error inserting template: " + ex.Message);
            }
        }




        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEnable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDisable;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReplaceAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHighlightAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuTemplates;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDoLetter;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSignalFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStatementOfCase;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStatementOfCaseDPM;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnServicePaper;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAgendaPts;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMoMeeting;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEmailFormat;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTourNotes;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAppreciation;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCaseStudy;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOpOrder;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnWarningOrder;
        internal RibbonButton button1;
        internal RibbonGroup group1;
        internal RibbonGroup group3;
        internal RibbonGroup group4;
        internal RibbonButton button2;
        internal RibbonButton highLightLike;
        internal RibbonSeparator separator1;
        internal RibbonSeparator separator2;
        internal RibbonSeparator separator3;
        internal RibbonSeparator separator4;
        internal RibbonMenu GovtService;
        internal RibbonMenu StaffPaper;
        internal RibbonButton btnGovtofindialetters;
        internal RibbonButton btnGovtofindia;
        internal RibbonButton btnServiceletters;
        internal RibbonButton btnNoteSheet;
        internal RibbonButton btnCopp;
        internal RibbonButton btnProblemStatement;
        internal RibbonMenu menu1;
        internal RibbonButton btnBrief;
        internal RibbonButton btnReturnBrief;
        internal RibbonButton btnNotice;
        internal RibbonButton btnCabinetNote;
        internal RibbonButton btnPressRelease;
        internal RibbonButton btnSocialMediaPost;
        internal RibbonButton btnGazetteNotification;
        internal RibbonMenu menu2;
        internal RibbonButton btnParliamentaryQuestionForwarding;
        internal RibbonButton btnParliamentaryQuestionReply;
        internal RibbonButton btnCoveringletterReply;
        internal RibbonMenu menu3;
        internal RibbonButton btnSampleWarningOrder;
        internal RibbonMenu menu4;
        internal RibbonButton btnMOMemorandum;
        internal RibbonButton btnBookReview;
        internal RibbonButton btnAnnotationinCorrespondence;
        public RibbonTab JSSD;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab JSSD_New;
    }

    partial class ThisRibbonCollection
    {
        internal AbbreviationRibbon AbbreviationRibbon
        {
            get { return this.GetRibbon<AbbreviationRibbon>(); }
        }
    }
}
