using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CoolTool
{
    public partial class MainWindow : Form
    {

        Modes mode = Modes.Logfile;

        public MainWindow()
        {
            InitializeComponent();
            filePath.SelectAll();
            if (Program.client == Client.Afford)
            {
                mode = Modes.Logfile;
                tabExcelSplit.Visible = false;
                tabFolderList.Visible = false;
                tabLangCopy.Visible = false;
                tabSlashing.Visible = false;
                tabTBmgmt.Visible = false;
                tabZaradek.Visible = false;
                tabTMList.Visible = false;
                tabPDF2DOC.Visible = false;
                tabSignatures.Visible = false;
                tabZaradekFinish.Visible = false;
                tabXTRFReport.Visible = false;
            }
            else if (Program.client == Client.Edimart)
            {
                mode = Modes.Logfile;
                tabLogFile.Select();
            }
            SetVisibility();
            Log.isProgress = true;
        }

        private void Browse_Click(object sender, EventArgs e)
        {
            if (mode == Modes.Zaradekolas || mode == Modes.Logfile
                    || mode == Modes.ExcelSplitter || mode == Modes.Slashing
                    || mode == Modes.Signature || mode == Modes.Signature || mode == Modes.XTRFReport)
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.CheckFileExists = true;
                DialogResult result = ofd.ShowDialog();
                if (result == DialogResult.OK)
                {
                    string file = ofd.FileName;

                    try
                    {
                        if (File.Exists(file))
                        {
                            filePath.Text = file;
                        }
                        else
                        {
                            throw new FileLoadException("File does not exist.");
                        }

                    }
                    catch (Exception ex)
                    {
                        Log.AddLog("File open error: " + ex.Message + "\r\n" + ex.StackTrace, true);
                    }
                }
            }
            else if (mode == Modes.TMList)
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.CheckFileExists = false;
                ofd.CheckPathExists = false;
                DialogResult result = ofd.ShowDialog();
                if (result == DialogResult.OK)
                {
                    string file = ofd.FileName;

                    try
                    {
                        if (Directory.Exists(Path.GetDirectoryName(file)) && Path.GetExtension(file).ToLower() == ".xlsx")
                        {
                            filePath.Text = file;
                        }
                        else
                        {
                            throw new FileLoadException("Folder does not exist or extension is not xlsx.");
                        }

                    }
                    catch (Exception ex)
                    {
                        Log.AddLog("File open error: " + ex.Message + "\r\n" + ex.StackTrace, true);
                    }
                }
            }
            else if (mode == Modes.LangCopier || mode == Modes.FolderList || mode == Modes.PDF2DOC || mode == Modes.ZaradekFinish)
            {

                FolderBrowserDialog ofd = new FolderBrowserDialog();
                DialogResult result = ofd.ShowDialog();
                if (result == DialogResult.OK)
                {
                    string file = ofd.SelectedPath;

                    try
                    {
                        if (Directory.Exists(file))
                        {
                            filePath.Text = file;
                        }
                        else
                        {
                            throw new FileLoadException("Folder does not exist.");
                        }

                    }
                    catch (Exception ex)
                    {
                        Log.AddLog("File open error: " + ex.Message + "\r\n" + ex.StackTrace, true);
                    }
                }
            }

        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            this.Dispose();
            return;
        }

        private void Convert_Click(object sender, EventArgs e)
        {
            modeChooser();
        }

        private void modeChooser()
        {
            try
            {
                updateProgress(0);
                string file = filePath.Text;
                if ((mode == Modes.Zaradekolas || mode == Modes.Logfile || mode == Modes.ExcelSplitter ||
                    mode == Modes.Slashing || mode == Modes.Signature || mode == Modes.XTRFReport) && !File.Exists(file))
                {
                    throw new FileLoadException("File does not exist.");
                }
                else if ((mode == Modes.LangCopier || mode == Modes.FolderList || mode == Modes.PDF2DOC) && !Directory.Exists(file))
                {
                    throw new FileLoadException("Folder does not exist.");
                }
                else if ((mode == Modes.TMList) && (!Directory.Exists(Path.GetDirectoryName(file)) || Path.GetExtension(file).ToLower() != ".xlsx"))
                {
                    throw new FileLoadException("Folder does not exist or file extension is not .xlsx");
                }

                switch (mode)
                {
                    case Modes.Zaradekolas:
                        Zaradek zaradekolo = new Zaradek(file, cbWMText.Text);
                        break;
                    case Modes.Logfile:
                        LogFile logfile = new LogFile(file, cbWCGrid.SelectedIndex);
                        break;
                    case Modes.LangCopier:
                        LangCopier langCopier = new LangCopier(file, clbLanguageList.CheckedItems);
                        break;
                    case Modes.FolderList:
                        FolderList folderList = new FolderList(file);
                        break;
                    case Modes.ExcelSplitter:
                        int excelMode = 0;
                        if (JoinExcel.Checked) excelMode = 1;
                        else if (SplitExcel.Checked) excelMode = 2;
                        else if (SavePDFExcel.Checked) excelMode = 3;
                        else if (styleChangerExcel.Checked) excelMode = 4;

                        if (excelMode != 0)
                        {
                            ExcelSplitter excelSplitter = new ExcelSplitter(file, excelMode);
                        }
                        break;
                    case Modes.PDF2DOC:
                        PDFHelper pdfHelper = new PDFHelper(file);
                        break;
                    case Modes.TMList:
                        TMList tmList = new TMList(file, tbTMListClient.Text, cbTMListType.SelectedIndex);
                        break;
                    case Modes.Slashing:
                        Slashing slashing = new Slashing(file, tbSlashing.Text);
                        break;
                    case Modes.TBmgmt:
                        TBmgmt tbmgmt = new TBmgmt(tbMgmtSubst.Text, tbMgmtReplace.Text, (Guid)cbTBMgmtTB.SelectedValue, cbTBMgmtCS.SelectedIndex, cbTBMgmtMatching.SelectedIndex);
                        break;
                    case Modes.Signature:
                        SignatureGenerator sign = new SignatureGenerator(file, tbSignaturesENText.Text, tbSignaturesHUText.Text);
                        break;
                    case Modes.ZaradekFinish:
                        int stampMode = 0;
                        if (radZaradekFinishLeft.Checked) stampMode = 1;
                        else if (radZaradekFinishRight.Checked) stampMode = 2;
                        else if (radZaradekFinishNostamp.Checked) stampMode = 3;
                        if (stampMode != 0)
                        {
                            ZaradekFinisher zf = new ZaradekFinisher(file, stampMode);
                        }
                        break;
                    case Modes.XTRFReport:
                        double exchangeRate;
                        if (!Double.TryParse(tbExchangeRate.Text.Replace(".", ","), out exchangeRate))
                        {
                            throw new Exception("Az árfolyam nem érvényes.");
                        }

                        XTRFGenerator xrep = new XTRFGenerator(file, exchangeRate);
                        break;
                }

            }
            catch (Exception ex)
            {
                Log.AddLog("Unexpected error " + ex.Message, true);
            }
            return;
        }

        public void updateProgress(int keszultseg)
        {
            if (keszultseg > 100)
            {
                keszultseg = 100;
            }
            pbKesz.Value = keszultseg;
            pbKesz.Refresh();
            Application.DoEvents();
        }

        private void tcModeChooser_SelectedChanged(object sender, TabControlEventArgs e)
        {
            if (tcModeChooser.SelectedTab == tabLogFile)
            {
                mode = Modes.Logfile;
            }
            else if (tcModeChooser.SelectedTab == tabFolderList)
            {
                mode = Modes.FolderList;
            }
            else if (tcModeChooser.SelectedTab == tabZaradek)
            {
                mode = Modes.Zaradekolas;
            }
            else if (tcModeChooser.SelectedTab == tabLangCopy)
            {
                mode = Modes.LangCopier;
            }
            else if (tcModeChooser.SelectedTab == tabExcelSplit)
            {
                mode = Modes.ExcelSplitter;
            }
            else if (tcModeChooser.SelectedTab == tabTMList)
            {
                mode = Modes.TMList;
            }
            else if (tcModeChooser.SelectedTab == tabPDF2DOC)
            {
                mode = Modes.PDF2DOC;
            }
            else if (tcModeChooser.SelectedTab == tabSlashing)
            {
                mode = Modes.Slashing;
            }
            else if (tcModeChooser.SelectedTab == tabTBmgmt)
            {
                mode = Modes.TBmgmt;
            }
            else if (tcModeChooser.SelectedTab == tabSignatures)
            {
                mode = Modes.Signature;
            }
            else if (tcModeChooser.SelectedTab == tabZaradekFinish)
            {
                mode = Modes.ZaradekFinish;
            }
            else if (tcModeChooser.SelectedTab == tabXTRFReport)
            {
                mode = Modes.XTRFReport;
            }

            SetVisibility();
        }


        private void SetVisibility()
        {
            try
            {
				Version version = Assembly.GetEntryAssembly().GetName().Version;
				lblVersion.Text = "v" + version.Major + "." + version.Minor + "." + version.Build;

				switch (mode)
                {
                    case Modes.Zaradekolas:
                        filePath.Enabled = true;
                        lblFilePath.Text = "Add meg a fájlt";
                        string[] WMOptions = new string[] { "TRANSLATION", "FORDÍTÁS", "ÜBERSETZUNG" };
                        cbWMText.DataSource = WMOptions;
                        cbWMText.DropDownStyle = ComboBoxStyle.DropDown;
                        break;
                    case Modes.Logfile:
                        filePath.Enabled = true;
                        lblFilePath.Text = "Add meg a fájlt";
                        List<string> gridnames = new List<string>();
                        foreach (string x in Program.grids.Keys)
                        {
                            gridnames.Add(x);
                        }
                        cbWCGrid.DataSource = gridnames;
                        break;
                    case Modes.LangCopier:
                        filePath.Enabled = true;
                        lblFilePath.Text = "Add meg a mappát";
                        clbLanguageList.Items.Clear();
                        clbLanguageList.Items.AddRange(LanguageHelper.LanguageNames());
                        break;
                    case Modes.FolderList:
                        filePath.Enabled = true;
                        lblFilePath.Text = "Add meg a mappát";
                        break;
                    case Modes.ExcelSplitter:
                        filePath.Enabled = true;
                        lblFilePath.Text = "Add meg a fájlt";
                        break;
                    case Modes.PDF2DOC:
                        filePath.Enabled = true;
                        lblFilePath.Text = "Add meg a mappát";
                        break;
                    case Modes.TMList:
                        filePath.Enabled = true;
                        lblFilePath.Text = "Add meg a fájlt";
                        WMOptions = new string[] { "Egyszerű lista", "Csoportosított lista" };
                        cbTMListType.DataSource = WMOptions;
                        cbTMListType.DropDownStyle = ComboBoxStyle.DropDown;
                        break;
                    case Modes.Slashing:
                        filePath.Enabled = true;
                        lblFilePath.Text = "Add meg a fájlt";
                        break;
                    case Modes.TBmgmt:
                        filePath.Enabled = false;
                        lblFilePath.Text = "Add meg a fájlt";
                        Dictionary<Guid, string> dataSource;
                        dataSource = Helper.TBList2stringArray(RAPI_Session.generateTBlist());
                        cbTBMgmtTB.DataSource = new BindingSource(dataSource, null);
                        cbTBMgmtTB.ValueMember = "Key";
                        cbTBMgmtTB.DisplayMember = "Value";
                        cbTBMgmtTB.DropDownStyle = ComboBoxStyle.DropDown;
						cbTBMgmtMatching.SelectedIndex = 1;
						cbTBMgmtCS.SelectedIndex = 1;
						break;
                    case Modes.Signature:
                        filePath.Enabled = true;
                        lblFilePath.Text = "Add meg a fájlt";
                        break;
                    case Modes.ZaradekFinish:
                        filePath.Enabled = true;
                        lblFilePath.Text = "Add meg a mappát";
                        break;
                    case Modes.XTRFReport:
                        filePath.Enabled = true;
                        lblFilePath.Text = "Add meg a fájlt";
                        break;

                }
            }
            catch (Exception ex)
            {
                Log.AddLog("UI initialization error: " + ex.Message, true);
            }
        }

        private void MainWindow_Load(object sender, EventArgs e)
        {

        }

        public void AddMessage(string message, bool type)
        {
            ProgressWindow.SelectionColor = type ? Color.Red : Color.Black;
            ProgressWindow.SelectionFont = new Font(ProgressWindow.SelectionFont, type ? FontStyle.Bold : FontStyle.Regular);
            ProgressWindow.AppendText(message);
            ProgressWindow.Refresh();
            ProgressWindow.SelectionStart = ProgressWindow.Text.Length;
            ProgressWindow.ScrollToCaret();
            Application.DoEvents();
        }

        private void ClearProgressWindow()
        {
            Log.AddLog("---------- Process completed ------------", true);
        }

    }
}
