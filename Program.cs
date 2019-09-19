using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aspose.Words;


namespace CoolTool
{
    static class Program
    {
        public static Client client = Client.Edimart;
        public static Dictionary<string, int[]> grids = new Dictionary<string, int[]>();
        public static int[] ColumnWidth;
        internal static int[] isCondForm;
        public static MainWindow mainWindow;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Aspose.Words.License Wlicense = new Aspose.Words.License();
            Wlicense.SetLicense("CoolTool.Aspose.Total.lic");
            Aspose.Cells.License XLlicense = new Aspose.Cells.License();
            XLlicense.SetLicense("CoolTool.Aspose.Total.lic");
            Aspose.Pdf.License Plicense = new Aspose.Pdf.License();
            Plicense.SetLicense("CoolTool.Aspose.Total.lic");
            
            if (client == Client.Edimart)
            {
                grids.Add("Vendor grid", new int[] { 15, 15, 5, 15, 30, 70, 100, 100, 100 });
                grids.Add("Ügyfél grid", new int[] { 30, 30, 20, 30, 50, 70, 100, 100, 100 });
                
            }
            else if (client == Client.Afford)
            {
                grids.Add("Vendor grid", new int[] { 0, 10, 10, 10, 25, 40, 65, 100, 100 });
                grids.Add("Ügyfél grid", new int[] { 0, 30, 30, 30, 45, 45, 75, 100, 100 });
                
            }

            LanguageHelper.Initialize();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            mainWindow = new MainWindow();

            Application.Run(mainWindow);
        }

    }

    public enum Modes
    {
        Zaradekolas,
        Logfile,
        LangCopier,
        FolderList,
        ExcelSplitter,
        PDF2DOC,
        TMList,
        Slashing,
        TBmgmt,
        Signature,
        ZaradekFinish,
        XTRFReport
    }

    public enum FileTypes
    {
        memoQ,
        memoQAllInfo,
        memoQTrados,
        memoQHTML,
        StudioXML,
        StudioHTML
    }

    public enum Client
    {
        Edimart,
        Afford
    }

}
