using System;
using System.Windows.Forms;
namespace CoolTool
{
    partial class MainWindow
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainWindow));
			this.lblFilePath = new System.Windows.Forms.Label();
			this.filePath = new System.Windows.Forms.TextBox();
			this.Browse = new System.Windows.Forms.Button();
			this.Convert = new System.Windows.Forms.Button();
			this.btnClose = new System.Windows.Forms.Button();
			this.cbWCGrid = new System.Windows.Forms.ComboBox();
			this.label1 = new System.Windows.Forms.Label();
			this.pbKesz = new System.Windows.Forms.ProgressBar();
			this.label3 = new System.Windows.Forms.Label();
			this.clbLanguageList = new System.Windows.Forms.CheckedListBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.styleChangerExcel = new System.Windows.Forms.RadioButton();
			this.SavePDFExcel = new System.Windows.Forms.RadioButton();
			this.SplitExcel = new System.Windows.Forms.RadioButton();
			this.JoinExcel = new System.Windows.Forms.RadioButton();
			this.tbTMListClient = new System.Windows.Forms.TextBox();
			this.label4 = new System.Windows.Forms.Label();
			this.ProgressWindow = new System.Windows.Forms.RichTextBox();
			this.label5 = new System.Windows.Forms.Label();
			this.tbSlashing = new System.Windows.Forms.TextBox();
			this.tcModeChooser = new System.Windows.Forms.TabControl();
			this.tabLogFile = new System.Windows.Forms.TabPage();
			this.tabZaradek = new System.Windows.Forms.TabPage();
			this.label6 = new System.Windows.Forms.Label();
			this.cbWMText = new System.Windows.Forms.ComboBox();
			this.tabZaradekFinish = new System.Windows.Forms.TabPage();
			this.radZaradekFinishNostamp = new System.Windows.Forms.RadioButton();
			this.label2 = new System.Windows.Forms.Label();
			this.radZaradekFinishRight = new System.Windows.Forms.RadioButton();
			this.radZaradekFinishLeft = new System.Windows.Forms.RadioButton();
			this.tabLangCopy = new System.Windows.Forms.TabPage();
			this.tabFolderList = new System.Windows.Forms.TabPage();
			this.tabExcelSplit = new System.Windows.Forms.TabPage();
			this.tabTMList = new System.Windows.Forms.TabPage();
			this.label9 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.cbTMListType = new System.Windows.Forms.ComboBox();
			this.tabPDF2DOC = new System.Windows.Forms.TabPage();
			this.tabSlashing = new System.Windows.Forms.TabPage();
			this.label10 = new System.Windows.Forms.Label();
			this.tabTBmgmt = new System.Windows.Forms.TabPage();
			this.label18 = new System.Windows.Forms.Label();
			this.label17 = new System.Windows.Forms.Label();
			this.tbMgmtReplace = new System.Windows.Forms.TextBox();
			this.label16 = new System.Windows.Forms.Label();
			this.cbTBMgmtMatching = new System.Windows.Forms.ComboBox();
			this.label15 = new System.Windows.Forms.Label();
			this.cbTBMgmtCS = new System.Windows.Forms.ComboBox();
			this.label12 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.cbTBMgmtTB = new System.Windows.Forms.ComboBox();
			this.tbMgmtSubst = new System.Windows.Forms.TextBox();
			this.tabSignatures = new System.Windows.Forms.TabPage();
			this.label13 = new System.Windows.Forms.Label();
			this.tbSignaturesHUText = new System.Windows.Forms.TextBox();
			this.label7 = new System.Windows.Forms.Label();
			this.tbSignaturesENText = new System.Windows.Forms.TextBox();
			this.tabXTRFReport = new System.Windows.Forms.TabPage();
			this.label14 = new System.Windows.Forms.Label();
			this.tbExchangeRate = new System.Windows.Forms.TextBox();
			this.lblVersion = new System.Windows.Forms.Label();
			this.groupBox2.SuspendLayout();
			this.tcModeChooser.SuspendLayout();
			this.tabLogFile.SuspendLayout();
			this.tabZaradek.SuspendLayout();
			this.tabZaradekFinish.SuspendLayout();
			this.tabLangCopy.SuspendLayout();
			this.tabExcelSplit.SuspendLayout();
			this.tabTMList.SuspendLayout();
			this.tabSlashing.SuspendLayout();
			this.tabTBmgmt.SuspendLayout();
			this.tabSignatures.SuspendLayout();
			this.tabXTRFReport.SuspendLayout();
			this.SuspendLayout();
			// 
			// lblFilePath
			// 
			this.lblFilePath.AutoSize = true;
			this.lblFilePath.Location = new System.Drawing.Point(22, 33);
			this.lblFilePath.Name = "lblFilePath";
			this.lblFilePath.Size = new System.Drawing.Size(80, 13);
			this.lblFilePath.TabIndex = 0;
			this.lblFilePath.Text = "Add meg a fájlt:";
			// 
			// filePath
			// 
			this.filePath.Location = new System.Drawing.Point(157, 33);
			this.filePath.Name = "filePath";
			this.filePath.Size = new System.Drawing.Size(661, 20);
			this.filePath.TabIndex = 1;
			this.filePath.Text = "Ide írd be az útvonalat...";
			// 
			// Browse
			// 
			this.Browse.Location = new System.Drawing.Point(825, 33);
			this.Browse.Name = "Browse";
			this.Browse.Size = new System.Drawing.Size(75, 23);
			this.Browse.TabIndex = 2;
			this.Browse.Text = "Tallózás";
			this.Browse.UseVisualStyleBackColor = true;
			this.Browse.Click += new System.EventHandler(this.Browse_Click);
			// 
			// Convert
			// 
			this.Convert.Location = new System.Drawing.Point(799, 351);
			this.Convert.Name = "Convert";
			this.Convert.Size = new System.Drawing.Size(101, 44);
			this.Convert.TabIndex = 8;
			this.Convert.Text = "Végrehajtás";
			this.Convert.UseVisualStyleBackColor = true;
			this.Convert.Click += new System.EventHandler(this.Convert_Click);
			// 
			// btnClose
			// 
			this.btnClose.Location = new System.Drawing.Point(825, 730);
			this.btnClose.Name = "btnClose";
			this.btnClose.Size = new System.Drawing.Size(75, 23);
			this.btnClose.TabIndex = 9;
			this.btnClose.Text = "Bezárás";
			this.btnClose.UseVisualStyleBackColor = true;
			this.btnClose.Click += new System.EventHandler(this.btnOK_Click);
			// 
			// cbWCGrid
			// 
			this.cbWCGrid.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbWCGrid.FormattingEnabled = true;
			this.cbWCGrid.Location = new System.Drawing.Point(138, 28);
			this.cbWCGrid.Name = "cbWCGrid";
			this.cbWCGrid.Size = new System.Drawing.Size(394, 21);
			this.cbWCGrid.TabIndex = 7;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(6, 31);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(93, 13);
			this.label1.TabIndex = 9;
			this.label1.Text = "Alkalmazandó grid";
			// 
			// pbKesz
			// 
			this.pbKesz.Location = new System.Drawing.Point(167, 400);
			this.pbKesz.Name = "pbKesz";
			this.pbKesz.Size = new System.Drawing.Size(733, 23);
			this.pbKesz.TabIndex = 8;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(29, 406);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(48, 13);
			this.label3.TabIndex = 7;
			this.label3.Text = "Progress";
			// 
			// clbLanguageList
			// 
			this.clbLanguageList.CheckOnClick = true;
			this.clbLanguageList.FormattingEnabled = true;
			this.clbLanguageList.Location = new System.Drawing.Point(9, 25);
			this.clbLanguageList.MultiColumn = true;
			this.clbLanguageList.Name = "clbLanguageList";
			this.clbLanguageList.Size = new System.Drawing.Size(852, 214);
			this.clbLanguageList.TabIndex = 6;
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.styleChangerExcel);
			this.groupBox2.Controls.Add(this.SavePDFExcel);
			this.groupBox2.Controls.Add(this.SplitExcel);
			this.groupBox2.Controls.Add(this.JoinExcel);
			this.groupBox2.Location = new System.Drawing.Point(6, 15);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(410, 122);
			this.groupBox2.TabIndex = 11;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Válaszd ki a munkafolyamat lépést";
			// 
			// styleChangerExcel
			// 
			this.styleChangerExcel.AutoSize = true;
			this.styleChangerExcel.Location = new System.Drawing.Point(17, 92);
			this.styleChangerExcel.Name = "styleChangerExcel";
			this.styleChangerExcel.Size = new System.Drawing.Size(175, 17);
			this.styleChangerExcel.TabIndex = 3;
			this.styleChangerExcel.TabStop = true;
			this.styleChangerExcel.Text = "Munkalapok formázása a\'la Edit";
			this.styleChangerExcel.UseVisualStyleBackColor = true;
			this.styleChangerExcel.Visible = false;
			// 
			// SavePDFExcel
			// 
			this.SavePDFExcel.AutoSize = true;
			this.SavePDFExcel.Location = new System.Drawing.Point(17, 69);
			this.SavePDFExcel.Name = "SavePDFExcel";
			this.SavePDFExcel.Size = new System.Drawing.Size(175, 17);
			this.SavePDFExcel.TabIndex = 2;
			this.SavePDFExcel.TabStop = true;
			this.SavePDFExcel.Text = "Munkalapok mentése PDF-ként";
			this.SavePDFExcel.UseVisualStyleBackColor = true;
			// 
			// SplitExcel
			// 
			this.SplitExcel.AutoSize = true;
			this.SplitExcel.Checked = true;
			this.SplitExcel.Location = new System.Drawing.Point(17, 22);
			this.SplitExcel.Name = "SplitExcel";
			this.SplitExcel.Size = new System.Drawing.Size(212, 17);
			this.SplitExcel.TabIndex = 1;
			this.SplitExcel.TabStop = true;
			this.SplitExcel.Text = "Excel munkalapok mentése külön fájlba";
			this.SplitExcel.UseVisualStyleBackColor = true;
			// 
			// JoinExcel
			// 
			this.JoinExcel.AutoSize = true;
			this.JoinExcel.Location = new System.Drawing.Point(17, 45);
			this.JoinExcel.Name = "JoinExcel";
			this.JoinExcel.Size = new System.Drawing.Size(173, 17);
			this.JoinExcel.TabIndex = 0;
			this.JoinExcel.TabStop = true;
			this.JoinExcel.Text = "Excel munkalapok összefűzése";
			this.JoinExcel.UseVisualStyleBackColor = true;
			// 
			// tbTMListClient
			// 
			this.tbTMListClient.Location = new System.Drawing.Point(149, 29);
			this.tbTMListClient.Name = "tbTMListClient";
			this.tbTMListClient.Size = new System.Drawing.Size(348, 20);
			this.tbTMListClient.TabIndex = 1;
			this.tbTMListClient.Text = "***All***";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(6, 7);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(59, 13);
			this.label4.TabIndex = 9;
			this.label4.Text = "Célnyelvek";
			// 
			// ProgressWindow
			// 
			this.ProgressWindow.Location = new System.Drawing.Point(32, 462);
			this.ProgressWindow.Name = "ProgressWindow";
			this.ProgressWindow.Size = new System.Drawing.Size(868, 262);
			this.ProgressWindow.TabIndex = 13;
			this.ProgressWindow.Text = "";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(29, 436);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(65, 13);
			this.label5.TabIndex = 12;
			this.label5.Text = "Progress log";
			// 
			// tbSlashing
			// 
			this.tbSlashing.Location = new System.Drawing.Point(156, 33);
			this.tbSlashing.Name = "tbSlashing";
			this.tbSlashing.Size = new System.Drawing.Size(240, 20);
			this.tbSlashing.TabIndex = 14;
			this.tbSlashing.Text = "¤";
			// 
			// tcModeChooser
			// 
			this.tcModeChooser.Controls.Add(this.tabLogFile);
			this.tcModeChooser.Controls.Add(this.tabZaradek);
			this.tcModeChooser.Controls.Add(this.tabZaradekFinish);
			this.tcModeChooser.Controls.Add(this.tabLangCopy);
			this.tcModeChooser.Controls.Add(this.tabFolderList);
			this.tcModeChooser.Controls.Add(this.tabExcelSplit);
			this.tcModeChooser.Controls.Add(this.tabTMList);
			this.tcModeChooser.Controls.Add(this.tabPDF2DOC);
			this.tcModeChooser.Controls.Add(this.tabSlashing);
			this.tcModeChooser.Controls.Add(this.tabTBmgmt);
			this.tcModeChooser.Controls.Add(this.tabSignatures);
			this.tcModeChooser.Controls.Add(this.tabXTRFReport);
			this.tcModeChooser.Location = new System.Drawing.Point(25, 73);
			this.tcModeChooser.Multiline = true;
			this.tcModeChooser.Name = "tcModeChooser";
			this.tcModeChooser.SelectedIndex = 0;
			this.tcModeChooser.Size = new System.Drawing.Size(875, 272);
			this.tcModeChooser.TabIndex = 15;
			this.tcModeChooser.Selected += new System.Windows.Forms.TabControlEventHandler(this.tcModeChooser_SelectedChanged);
			// 
			// tabLogFile
			// 
			this.tabLogFile.Controls.Add(this.cbWCGrid);
			this.tabLogFile.Controls.Add(this.label1);
			this.tabLogFile.Location = new System.Drawing.Point(4, 40);
			this.tabLogFile.Name = "tabLogFile";
			this.tabLogFile.Padding = new System.Windows.Forms.Padding(3);
			this.tabLogFile.Size = new System.Drawing.Size(867, 228);
			this.tabLogFile.TabIndex = 0;
			this.tabLogFile.Text = "Szószám logfájl";
			this.tabLogFile.UseVisualStyleBackColor = true;
			// 
			// tabZaradek
			// 
			this.tabZaradek.Controls.Add(this.label6);
			this.tabZaradek.Controls.Add(this.cbWMText);
			this.tabZaradek.Location = new System.Drawing.Point(4, 40);
			this.tabZaradek.Name = "tabZaradek";
			this.tabZaradek.Padding = new System.Windows.Forms.Padding(3);
			this.tabZaradek.Size = new System.Drawing.Size(867, 228);
			this.tabZaradek.TabIndex = 4;
			this.tabZaradek.Text = "Záradék előkészítése";
			this.tabZaradek.UseVisualStyleBackColor = true;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(6, 33);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(76, 13);
			this.label6.TabIndex = 9;
			this.label6.Text = "Vízjel szövege";
			// 
			// cbWMText
			// 
			this.cbWMText.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbWMText.FormattingEnabled = true;
			this.cbWMText.Location = new System.Drawing.Point(128, 30);
			this.cbWMText.Name = "cbWMText";
			this.cbWMText.Size = new System.Drawing.Size(328, 21);
			this.cbWMText.TabIndex = 7;
			// 
			// tabZaradekFinish
			// 
			this.tabZaradekFinish.Controls.Add(this.radZaradekFinishNostamp);
			this.tabZaradekFinish.Controls.Add(this.label2);
			this.tabZaradekFinish.Controls.Add(this.radZaradekFinishRight);
			this.tabZaradekFinish.Controls.Add(this.radZaradekFinishLeft);
			this.tabZaradekFinish.Location = new System.Drawing.Point(4, 40);
			this.tabZaradekFinish.Name = "tabZaradekFinish";
			this.tabZaradekFinish.Padding = new System.Windows.Forms.Padding(3);
			this.tabZaradekFinish.Size = new System.Drawing.Size(867, 228);
			this.tabZaradekFinish.TabIndex = 10;
			this.tabZaradekFinish.Text = "Záradék összefűzése";
			this.tabZaradekFinish.UseVisualStyleBackColor = true;
			// 
			// radZaradekFinishNostamp
			// 
			this.radZaradekFinishNostamp.AutoSize = true;
			this.radZaradekFinishNostamp.Location = new System.Drawing.Point(47, 80);
			this.radZaradekFinishNostamp.Name = "radZaradekFinishNostamp";
			this.radZaradekFinishNostamp.Size = new System.Drawing.Size(87, 17);
			this.radZaradekFinishNostamp.TabIndex = 3;
			this.radZaradekFinishNostamp.Text = "Nincs pecsét";
			this.radZaradekFinishNostamp.UseVisualStyleBackColor = true;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(18, 13);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(85, 13);
			this.label2.TabIndex = 2;
			this.label2.Text = "Pecsét helyzete:";
			// 
			// radZaradekFinishRight
			// 
			this.radZaradekFinishRight.AutoSize = true;
			this.radZaradekFinishRight.Location = new System.Drawing.Point(47, 56);
			this.radZaradekFinishRight.Name = "radZaradekFinishRight";
			this.radZaradekFinishRight.Size = new System.Drawing.Size(85, 17);
			this.radZaradekFinishRight.TabIndex = 1;
			this.radZaradekFinishRight.Text = "Jobb oldalon";
			this.radZaradekFinishRight.UseVisualStyleBackColor = true;
			// 
			// radZaradekFinishLeft
			// 
			this.radZaradekFinishLeft.AutoSize = true;
			this.radZaradekFinishLeft.Checked = true;
			this.radZaradekFinishLeft.Location = new System.Drawing.Point(47, 32);
			this.radZaradekFinishLeft.Name = "radZaradekFinishLeft";
			this.radZaradekFinishLeft.Size = new System.Drawing.Size(77, 17);
			this.radZaradekFinishLeft.TabIndex = 0;
			this.radZaradekFinishLeft.TabStop = true;
			this.radZaradekFinishLeft.Text = "Bal oldalon";
			this.radZaradekFinishLeft.UseVisualStyleBackColor = true;
			// 
			// tabLangCopy
			// 
			this.tabLangCopy.Controls.Add(this.clbLanguageList);
			this.tabLangCopy.Controls.Add(this.label4);
			this.tabLangCopy.Location = new System.Drawing.Point(4, 40);
			this.tabLangCopy.Name = "tabLangCopy";
			this.tabLangCopy.Padding = new System.Windows.Forms.Padding(3);
			this.tabLangCopy.Size = new System.Drawing.Size(867, 228);
			this.tabLangCopy.TabIndex = 2;
			this.tabLangCopy.Text = "Nyelv másolása";
			this.tabLangCopy.UseVisualStyleBackColor = true;
			// 
			// tabFolderList
			// 
			this.tabFolderList.Location = new System.Drawing.Point(4, 40);
			this.tabFolderList.Name = "tabFolderList";
			this.tabFolderList.Padding = new System.Windows.Forms.Padding(3);
			this.tabFolderList.Size = new System.Drawing.Size(867, 228);
			this.tabFolderList.TabIndex = 1;
			this.tabFolderList.Text = "Fájllista";
			this.tabFolderList.UseVisualStyleBackColor = true;
			// 
			// tabExcelSplit
			// 
			this.tabExcelSplit.Controls.Add(this.groupBox2);
			this.tabExcelSplit.Location = new System.Drawing.Point(4, 40);
			this.tabExcelSplit.Name = "tabExcelSplit";
			this.tabExcelSplit.Padding = new System.Windows.Forms.Padding(3);
			this.tabExcelSplit.Size = new System.Drawing.Size(867, 228);
			this.tabExcelSplit.TabIndex = 3;
			this.tabExcelSplit.Text = "Exceldarabolás";
			this.tabExcelSplit.UseVisualStyleBackColor = true;
			// 
			// tabTMList
			// 
			this.tabTMList.Controls.Add(this.label9);
			this.tabTMList.Controls.Add(this.label8);
			this.tabTMList.Controls.Add(this.cbTMListType);
			this.tabTMList.Controls.Add(this.tbTMListClient);
			this.tabTMList.Location = new System.Drawing.Point(4, 40);
			this.tabTMList.Name = "tabTMList";
			this.tabTMList.Padding = new System.Windows.Forms.Padding(3);
			this.tabTMList.Size = new System.Drawing.Size(867, 228);
			this.tabTMList.TabIndex = 5;
			this.tabTMList.Text = "TM lista";
			this.tabTMList.UseVisualStyleBackColor = true;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(6, 78);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(65, 13);
			this.label9.TabIndex = 9;
			this.label9.Text = "Lista típusa:";
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(6, 32);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(106, 13);
			this.label8.TabIndex = 9;
			this.label8.Text = "Add meg az ügyfelet:";
			// 
			// cbTMListType
			// 
			this.cbTMListType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbTMListType.FormattingEnabled = true;
			this.cbTMListType.Location = new System.Drawing.Point(149, 75);
			this.cbTMListType.Name = "cbTMListType";
			this.cbTMListType.Size = new System.Drawing.Size(348, 21);
			this.cbTMListType.TabIndex = 7;
			// 
			// tabPDF2DOC
			// 
			this.tabPDF2DOC.Location = new System.Drawing.Point(4, 40);
			this.tabPDF2DOC.Name = "tabPDF2DOC";
			this.tabPDF2DOC.Padding = new System.Windows.Forms.Padding(3);
			this.tabPDF2DOC.Size = new System.Drawing.Size(867, 228);
			this.tabPDF2DOC.TabIndex = 6;
			this.tabPDF2DOC.Text = "PDF2DOC";
			this.tabPDF2DOC.UseVisualStyleBackColor = true;
			// 
			// tabSlashing
			// 
			this.tabSlashing.Controls.Add(this.label10);
			this.tabSlashing.Controls.Add(this.tbSlashing);
			this.tabSlashing.Location = new System.Drawing.Point(4, 40);
			this.tabSlashing.Name = "tabSlashing";
			this.tabSlashing.Padding = new System.Windows.Forms.Padding(3);
			this.tabSlashing.Size = new System.Drawing.Size(867, 228);
			this.tabSlashing.TabIndex = 7;
			this.tabSlashing.Text = "Perjelezés";
			this.tabSlashing.UseVisualStyleBackColor = true;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(6, 36);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(135, 13);
			this.label10.TabIndex = 9;
			this.label10.Text = "Add meg az elválasztójelet:";
			// 
			// tabTBmgmt
			// 
			this.tabTBmgmt.Controls.Add(this.label18);
			this.tabTBmgmt.Controls.Add(this.label17);
			this.tabTBmgmt.Controls.Add(this.tbMgmtReplace);
			this.tabTBmgmt.Controls.Add(this.label16);
			this.tabTBmgmt.Controls.Add(this.cbTBMgmtMatching);
			this.tabTBmgmt.Controls.Add(this.label15);
			this.tabTBmgmt.Controls.Add(this.cbTBMgmtCS);
			this.tabTBmgmt.Controls.Add(this.label12);
			this.tabTBmgmt.Controls.Add(this.label11);
			this.tabTBmgmt.Controls.Add(this.cbTBMgmtTB);
			this.tabTBmgmt.Controls.Add(this.tbMgmtSubst);
			this.tabTBmgmt.Location = new System.Drawing.Point(4, 40);
			this.tabTBmgmt.Name = "tabTBmgmt";
			this.tabTBmgmt.Padding = new System.Windows.Forms.Padding(3);
			this.tabTBmgmt.Size = new System.Drawing.Size(867, 228);
			this.tabTBmgmt.TabIndex = 8;
			this.tabTBmgmt.Text = "TB kezelés";
			this.tabTBmgmt.UseVisualStyleBackColor = true;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(6, 124);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(118, 13);
			this.label18.TabIndex = 21;
			this.label18.Text = "Alapértelmezett értékek";
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(390, 30);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(91, 13);
			this.label17.TabIndex = 19;
			this.label17.Text = "Cserélendő string:";
			// 
			// tbMgmtReplace
			// 
			this.tbMgmtReplace.Location = new System.Drawing.Point(498, 27);
			this.tbMgmtReplace.Name = "tbMgmtReplace";
			this.tbMgmtReplace.Size = new System.Drawing.Size(240, 20);
			this.tbMgmtReplace.TabIndex = 20;
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.Location = new System.Drawing.Point(6, 151);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(47, 13);
			this.label16.TabIndex = 18;
			this.label16.Text = "Illesztés:";
			// 
			// cbTBMgmtMatching
			// 
			this.cbTBMgmtMatching.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbTBMgmtMatching.FormattingEnabled = true;
			this.cbTBMgmtMatching.Items.AddRange(new object[] {
            "Nem pontos",
            "50%-os előtag",
            "Pontos",
            "Egyedi"});
			this.cbTBMgmtMatching.Location = new System.Drawing.Point(128, 148);
			this.cbTBMgmtMatching.Name = "cbTBMgmtMatching";
			this.cbTBMgmtMatching.Size = new System.Drawing.Size(172, 21);
			this.cbTBMgmtMatching.TabIndex = 17;
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.Location = new System.Drawing.Point(6, 185);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(119, 13);
			this.label15.TabIndex = 16;
			this.label15.Text = "Nagybetű-érzékenység:";
			// 
			// cbTBMgmtCS
			// 
			this.cbTBMgmtCS.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbTBMgmtCS.FormattingEnabled = true;
			this.cbTBMgmtCS.Items.AddRange(new object[] {
            "Igen",
            "Megengedő",
            "Nem"});
			this.cbTBMgmtCS.Location = new System.Drawing.Point(128, 182);
			this.cbTBMgmtCS.Name = "cbTBMgmtCS";
			this.cbTBMgmtCS.Size = new System.Drawing.Size(172, 21);
			this.cbTBMgmtCS.TabIndex = 15;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(6, 73);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(85, 13);
			this.label12.TabIndex = 9;
			this.label12.Text = "TB kiválasztása:";
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(6, 30);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(95, 13);
			this.label11.TabIndex = 9;
			this.label11.Text = "Helyettesítő string:";
			// 
			// cbTBMgmtTB
			// 
			this.cbTBMgmtTB.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbTBMgmtTB.FormattingEnabled = true;
			this.cbTBMgmtTB.Location = new System.Drawing.Point(128, 70);
			this.cbTBMgmtTB.Name = "cbTBMgmtTB";
			this.cbTBMgmtTB.Size = new System.Drawing.Size(455, 21);
			this.cbTBMgmtTB.TabIndex = 7;
			// 
			// tbMgmtSubst
			// 
			this.tbMgmtSubst.Location = new System.Drawing.Point(128, 27);
			this.tbMgmtSubst.Name = "tbMgmtSubst";
			this.tbMgmtSubst.Size = new System.Drawing.Size(240, 20);
			this.tbMgmtSubst.TabIndex = 14;
			this.tbMgmtSubst.Text = "///";
			// 
			// tabSignatures
			// 
			this.tabSignatures.Controls.Add(this.label13);
			this.tabSignatures.Controls.Add(this.tbSignaturesHUText);
			this.tabSignatures.Controls.Add(this.label7);
			this.tabSignatures.Controls.Add(this.tbSignaturesENText);
			this.tabSignatures.Location = new System.Drawing.Point(4, 40);
			this.tabSignatures.Name = "tabSignatures";
			this.tabSignatures.Padding = new System.Windows.Forms.Padding(3);
			this.tabSignatures.Size = new System.Drawing.Size(867, 228);
			this.tabSignatures.TabIndex = 9;
			this.tabSignatures.Text = "Aláírások létrehozása";
			this.tabSignatures.UseVisualStyleBackColor = true;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.Location = new System.Drawing.Point(6, 78);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(123, 13);
			this.label13.TabIndex = 9;
			this.label13.Text = "Magyar speciális üzenet:";
			// 
			// tbSignaturesHUText
			// 
			this.tbSignaturesHUText.Location = new System.Drawing.Point(138, 75);
			this.tbSignaturesHUText.Name = "tbSignaturesHUText";
			this.tbSignaturesHUText.Size = new System.Drawing.Size(723, 20);
			this.tbSignaturesHUText.TabIndex = 14;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(6, 36);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(115, 13);
			this.label7.TabIndex = 9;
			this.label7.Text = "Angol speciális üzenet:";
			// 
			// tbSignaturesENText
			// 
			this.tbSignaturesENText.Location = new System.Drawing.Point(138, 33);
			this.tbSignaturesENText.Name = "tbSignaturesENText";
			this.tbSignaturesENText.Size = new System.Drawing.Size(723, 20);
			this.tbSignaturesENText.TabIndex = 14;
			// 
			// tabXTRFReport
			// 
			this.tabXTRFReport.Controls.Add(this.label14);
			this.tabXTRFReport.Controls.Add(this.tbExchangeRate);
			this.tabXTRFReport.Location = new System.Drawing.Point(4, 40);
			this.tabXTRFReport.Name = "tabXTRFReport";
			this.tabXTRFReport.Padding = new System.Windows.Forms.Padding(3);
			this.tabXTRFReport.Size = new System.Drawing.Size(867, 228);
			this.tabXTRFReport.TabIndex = 11;
			this.tabXTRFReport.Text = "XTRF riport";
			this.tabXTRFReport.UseVisualStyleBackColor = true;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.Location = new System.Drawing.Point(19, 32);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(75, 13);
			this.label14.TabIndex = 11;
			this.label14.Text = "EUR árfolyam:";
			// 
			// tbExchangeRate
			// 
			this.tbExchangeRate.Location = new System.Drawing.Point(162, 29);
			this.tbExchangeRate.Name = "tbExchangeRate";
			this.tbExchangeRate.Size = new System.Drawing.Size(348, 20);
			this.tbExchangeRate.TabIndex = 10;
			this.tbExchangeRate.Text = "310";
			// 
			// lblVersion
			// 
			this.lblVersion.AutoSize = true;
			this.lblVersion.Font = new System.Drawing.Font("Microsoft Sans Serif", 6F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
			this.lblVersion.Location = new System.Drawing.Point(6, 748);
			this.lblVersion.Name = "lblVersion";
			this.lblVersion.Size = new System.Drawing.Size(0, 9);
			this.lblVersion.TabIndex = 16;
			this.lblVersion.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// MainWindow
			// 
			this.AcceptButton = this.Convert;
			this.ClientSize = new System.Drawing.Size(927, 761);
			this.Controls.Add(this.lblVersion);
			this.Controls.Add(this.tcModeChooser);
			this.Controls.Add(this.ProgressWindow);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.pbKesz);
			this.Controls.Add(this.btnClose);
			this.Controls.Add(this.Convert);
			this.Controls.Add(this.Browse);
			this.Controls.Add(this.filePath);
			this.Controls.Add(this.lblFilePath);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Name = "MainWindow";
			this.Load += new System.EventHandler(this.MainWindow_Load);
			this.groupBox2.ResumeLayout(false);
			this.groupBox2.PerformLayout();
			this.tcModeChooser.ResumeLayout(false);
			this.tabLogFile.ResumeLayout(false);
			this.tabLogFile.PerformLayout();
			this.tabZaradek.ResumeLayout(false);
			this.tabZaradek.PerformLayout();
			this.tabZaradekFinish.ResumeLayout(false);
			this.tabZaradekFinish.PerformLayout();
			this.tabLangCopy.ResumeLayout(false);
			this.tabLangCopy.PerformLayout();
			this.tabExcelSplit.ResumeLayout(false);
			this.tabTMList.ResumeLayout(false);
			this.tabTMList.PerformLayout();
			this.tabSlashing.ResumeLayout(false);
			this.tabSlashing.PerformLayout();
			this.tabTBmgmt.ResumeLayout(false);
			this.tabTBmgmt.PerformLayout();
			this.tabSignatures.ResumeLayout(false);
			this.tabSignatures.PerformLayout();
			this.tabXTRFReport.ResumeLayout(false);
			this.tabXTRFReport.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

        }


        #endregion

        private System.Windows.Forms.Label lblFilePath;
        private System.Windows.Forms.TextBox filePath;
        private System.Windows.Forms.Button Browse;
        private System.Windows.Forms.Button Convert;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.ComboBox cbWCGrid;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ProgressBar pbKesz;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckedListBox clbLanguageList;
        private GroupBox groupBox2;
        private RadioButton SplitExcel;
        private RadioButton JoinExcel;
        private TextBox tbTMListClient;
        private Label label4;
        private RadioButton SavePDFExcel;
        private RichTextBox ProgressWindow;
        private Label label5;
        private RadioButton styleChangerExcel;
        private TextBox tbSlashing;
        private TabControl tcModeChooser;
        private TabPage tabLogFile;
        private TabPage tabFolderList;
        private TabPage tabLangCopy;
        private TabPage tabExcelSplit;
        private TabPage tabZaradek;
        private ComboBox cbWMText;
        private Label label6;
        private TabPage tabTMList;
        private TabPage tabPDF2DOC;
        private TabPage tabSlashing;
        private TabPage tabTBmgmt;
        private TabPage tabSignatures;
        private Label label9;
        private Label label8;
        private ComboBox cbTMListType;
        private Label label10;
        private Label label12;
        private Label label11;
        private ComboBox cbTBMgmtTB;
        private TextBox tbMgmtSubst;
        private Label label13;
        private TextBox tbSignaturesHUText;
        private Label label7;
        private TextBox tbSignaturesENText;
        private TabPage tabZaradekFinish;
        private RadioButton radZaradekFinishNostamp;
        private Label label2;
        private RadioButton radZaradekFinishRight;
        private RadioButton radZaradekFinishLeft;
        private TabPage tabXTRFReport;
        private Label label14;
        private TextBox tbExchangeRate;
		private Label label17;
		private TextBox tbMgmtReplace;
		private Label label16;
		private ComboBox cbTBMgmtMatching;
		private Label label15;
		private ComboBox cbTBMgmtCS;
		private Label label18;
		private Label lblVersion;
	}
}

