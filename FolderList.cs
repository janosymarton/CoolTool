using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Cells;
using Aspose.Pdf;
using Aspose.Pdf.Facades;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells.Drawing;
using Aspose.Slides;
using System.Windows.Forms;

namespace CoolTool
{
    internal class FolderList
    {
        private string baseFolder;
        
        private List<FileObject> fileList = new List<FileObject>();
        private int progress = 0;
        private string tmpFolder = Path.GetTempPath();
        public FolderList(string file)
        {
            this.baseFolder = file;
            
            try
            {
                Program.mainWindow.updateProgress(0);
                RunFileCollector(baseFolder);
                Program.mainWindow.updateProgress(90);
                RunResultExport();
                Program.mainWindow.updateProgress(100);
                Log.AddLog("Processing finished");
            }
            catch (Exception ex)
            {
                Log.AddLog("Problem while processing: " + ex.Message, true);
            }

        }

        private void RunFileCollector(string currentFolder)
        {
            string[] filesInCurrentFolder = Directory.GetFiles(currentFolder, "*.*", SearchOption.AllDirectories);
            int countOfFiles = filesInCurrentFolder.Length;
            int fileCount = 0;
            
            foreach (string file in filesInCurrentFolder)
            {
                FileObject fo = new FileObject(file);
                fileList.Add(fo);

                try
                {
                    switch (Path.GetExtension(file).ToLower())
                    {
                        case ".zip":
                            int increment = Convert.ToInt32(80.0 / Convert.ToDouble(countOfFiles));
                            ParseZip(fo.fullPath, fo.fullPath, increment);
                            break;
                        case ".doc":
                        case ".docx":
                            ParseDoc(ref fo, fo.fullPath);
                            break;
                        case ".pdf":
                            ParsePDF(ref fo, fo.fullPath);
                            break;
                        case ".xls":
                        case ".xlsx":
                            ParseXls(ref fo, fo.fullPath);
                            break;
                        case ".ppt":
                        case ".pptx":
                            ParsePpt(ref fo, fo.fullPath);
                            break;

                    }
                }
                catch (Exception ex)
                {
                    fo.comment = "File parsing error: " + ex.Message;
                }

                fileCount++;
                progress = Convert.ToInt32( 80.0 * Convert.ToDouble(fileCount) / Convert.ToDouble(countOfFiles));
                Program.mainWindow.updateProgress(progress);
            }

        }

        private void ParsePpt(ref FileObject fo, string fullPath)
        {
            Presentation pres = new Presentation(fullPath);
            fo.pageCount = pres.Slides.Count;
            fo.hasPassword = pres.ProtectionManager.IsEncrypted || pres.ProtectionManager.IsWriteProtected;
        }

        private void ParseXls(ref FileObject fo, string filePath)
        {
            Workbook wb = new Workbook(filePath);
            Aspose.Cells.Properties.BuiltInDocumentPropertyCollection dp = wb.BuiltInDocumentProperties;
            WorksheetCollection wsc = wb.Worksheets;
            fo.pageCount = wsc.Count;

            int NoOfImages = 0;
            int NoOfEmbeddedDocs = 0;
            bool isProtected = false;
            foreach (Worksheet ws in wsc)
            {
                OleObjectCollection oles = ws.OleObjects;
                if (ws.IsProtected) isProtected = true;
                foreach (OleObject ole in oles)
                {
                    switch (ole.FileFormatType)
                    {
                        case FileFormatType.Doc:
                        case FileFormatType.Xlsm:
                        case FileFormatType.Docx:
                        case FileFormatType.Xlsx:
                        case FileFormatType.Ppt:
                        case FileFormatType.Pdf:
                        case FileFormatType.CSV:
                        case FileFormatType.VSD:
                        case FileFormatType.VSDX:
                        case FileFormatType.Html:
                        case FileFormatType.XML:
                            NoOfEmbeddedDocs++;
                            break;
                        case FileFormatType.BMP:
                        case FileFormatType.TIFF:
                            NoOfImages++;
                            break;
                        default:
                            NoOfImages++;
                            break;
                    }
                }

            }

            fo.embeddedDocsCount = NoOfEmbeddedDocs;
            fo.imageCount = NoOfImages;
            fo.hasPassword = isProtected;

            string tmpFolderToExtract = tmpFolder + "\\" + Guid.NewGuid();
            Directory.CreateDirectory(tmpFolderToExtract);
            string tmpTextFile = tmpFolderToExtract + "\\" + "tmpTextexport.txt";

            byte[] workbookData = new byte[0];
            TxtSaveOptions opts = new TxtSaveOptions();
            opts.Separator = ' ';

            for (int idx = 0; idx < wb.Worksheets.Count; idx++)
            {
                MemoryStream ms = new MemoryStream();
                wb.Worksheets.ActiveSheetIndex = idx;
                wb.Save(ms, opts);
                ms.Position = 0;
                byte[] sheetData = ms.ToArray();
                byte[] combinedArray = new byte[workbookData.Length + sheetData.Length];
                Array.Copy(workbookData, 0, combinedArray, 0, workbookData.Length);
                Array.Copy(sheetData, 0, combinedArray, workbookData.Length, sheetData.Length);
                workbookData = combinedArray;
            }

            File.WriteAllBytes(tmpTextFile, workbookData);

            fo.wordCount = GetWordCount(tmpTextFile);
            fo.characterCount = GetCharCount(tmpTextFile);
            if (File.Exists(tmpTextFile))
            {
                File.Delete(tmpTextFile);
            }
            if (Directory.Exists(tmpFolderToExtract))
            {
                Directory.Delete(tmpFolderToExtract);
            }

        }

        private void ParsePDF(ref FileObject fo, string filePath)
        {
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(filePath);
            PdfFileInfo pi = new PdfFileInfo(pdfDocument);
            PdfExtractor pe = new PdfExtractor(pdfDocument);
            ImagePlacementAbsorber abs = new ImagePlacementAbsorber();

            fo.pageCount = pi.NumberOfPages;
            fo.embeddedDocsCount = pdfDocument.EmbeddedFiles.Count;
            pdfDocument.Pages.Accept(abs);
            fo.imageCount = abs.ImagePlacements.Count;
            fo.hasPassword = pi.HasOpenPassword;
            pe.ExtractText(Encoding.ASCII);
            string tmpFolderToExtract = tmpFolder + "\\" + Guid.NewGuid();
            Directory.CreateDirectory(tmpFolderToExtract);
            string tmpTextFile = tmpFolderToExtract + "\\" + "tmpTextexport.txt";
            pe.GetText(tmpTextFile);
            fo.wordCount = GetWordCount(tmpTextFile);
            fo.characterCount = GetCharCount(tmpTextFile);
            if (File.Exists(tmpTextFile))
            {
                File.Delete(tmpTextFile);
            }
            if (Directory.Exists(tmpFolderToExtract))
            {
                Directory.Delete(tmpFolderToExtract);
            }

        }

        private int GetWordCount(string filePath)
        {
            int WC = 0;

            using (StreamReader sr = new StreamReader(filePath))
            {
                string line = "";
                while (!sr.EndOfStream)
                {
                    line = sr.ReadLine();
                    Regex letterRgx = new Regex(@"[\p{L}\d]");
                    bool bmc = letterRgx.IsMatch(line);
                    if (bmc)
                    {
                        Regex rgx = new Regex(@"[^\s](\s+)[^\s]");
                        MatchCollection mc = rgx.Matches(line);
                        WC += mc.Count + 1;
                    }
                }
            }

            return WC;
        }

        private int GetCharCount(string filePath)
        {
            int CC = 0;

            using (StreamReader sr = new StreamReader(filePath))
            {
                string line = "";
                while (!sr.EndOfStream)
                {
                    line = sr.ReadLine();
                    CC += line.Length;
                }
            }

            return CC;
        }


        private void ParseDoc(ref FileObject fo, string filePath)
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(filePath);
            Aspose.Words.Properties.BuiltInDocumentProperties dp = doc.BuiltInDocumentProperties;
            fo.wordCount = dp.Words;
            fo.characterCount = dp.CharactersWithSpaces;
            fo.pageCount = dp.Pages;
            fo.hasPassword = dp.Security == Aspose.Words.Properties.DocumentSecurity.PasswordProtected;

            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int NoOfImages = 0;
            foreach (Aspose.Words.Drawing.Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    NoOfImages++;
                }
            }
            fo.imageCount = NoOfImages;

            NodeCollection embeddedDocs = doc.GetChildNodes(NodeType.SubDocument, true);

            fo.embeddedDocsCount = embeddedDocs.Count;

        }

        private void ParseZip(string baseFolder, string filePath, int increment)
        {
            using (ZipArchive zipFile = ZipFile.OpenRead(filePath))
            {
                IReadOnlyCollection<ZipArchiveEntry> zippedFiles = zipFile.Entries;
                string tmpUnzipFolder = tmpFolder + "\\" + Guid.NewGuid();
                Directory.CreateDirectory(tmpUnzipFolder);
                int noOfFiles = zippedFiles.Count;
                int FileCount = 0;

                foreach (ZipArchiveEntry zippedFile in zippedFiles)
                {
                    string zippedPath = zippedFile.FullName;
                    if (zippedPath.Substring(zippedPath.Length - 1) != "/")
                    {
                        FileObject fo = new FileObject(baseFolder + "\\" + zippedPath.Replace('/', '\\'));
                        fileList.Add(fo);
                        string tmpFile = Path.Combine(tmpUnzipFolder, zippedFile.Name);
                        try
                        {

                            switch (Path.GetExtension(zippedPath).ToLower())
                            {
                                case ".zip":
                                    zippedFile.ExtractToFile(tmpFile, true);
                                    int zipincrement = Convert.ToInt32(Convert.ToDouble(increment) / Convert.ToDouble(noOfFiles));
                                    ParseZip(fo.fullPath, tmpFile, zipincrement);
                                    File.Delete(tmpFile);
                                    break;
                                case ".doc":
                                case ".docx":
                                    zippedFile.ExtractToFile(tmpFile, true);
                                    ParseDoc(ref fo, tmpFile);
                                    File.Delete(tmpFile);
                                    break;
                                case ".pdf":
                                    zippedFile.ExtractToFile(tmpFile, true);
                                    ParsePDF(ref fo, tmpFile);
                                    File.Delete(tmpFile);
                                    break;
                                case ".xls":
                                case ".xlsx":
                                    zippedFile.ExtractToFile(tmpFile, true);
                                    ParseXls(ref fo, tmpFile);
                                    File.Delete(tmpFile);
                                    break;
                                case ".ppt":
                                case ".pptx":
                                    zippedFile.ExtractToFile(tmpFile, true);
                                    ParsePpt(ref fo, tmpFile);
                                    File.Delete(tmpFile);
                                    break;
                            }
                        }
                        catch (Exception ex)
                        {
                            fo.comment = "File parsing error: " + ex.Message;
                        }
                        FileCount++;
                        progress = Convert.ToInt32((increment * Convert.ToDouble(FileCount) / Convert.ToDouble(noOfFiles)));
                        Program.mainWindow.updateProgress(progress);
                    }
                }
                Directory.Delete(tmpUnzipFolder);
            }
        }

        private void RunResultExport()
        {
            string newXLSPath = Path.Combine(baseFolder, "_FolderContent.xlsx");

            int i = 1;
            while (File.Exists(newXLSPath))
            {
                newXLSPath = Path.Combine(baseFolder, "_FolderContent" + i.ToString() + ".xlsx");
                i++;
            }

            Workbook wb = new Workbook();
            Worksheet ws;
            string wsName = "Results";
            if (wb.Worksheets.Count == 0)
            {
                ws = wb.Worksheets.Add(wsName);
            }
            else
            {
                ws = wb.Worksheets[0];
                ws.Name = wsName;
            }
            Aspose.Cells.Cell c;

            c = ws.Cells[0, 0];
            c.Value = "Full Path";

            c = ws.Cells[0, 1];
            c.Value = "Path";

            c = ws.Cells[0, 2];
            c.Value = "Filename";

            c = ws.Cells[0, 3];
            c.Value = "File type";

            c = ws.Cells[0, 4];
            c.Value = "Page count";

            c = ws.Cells[0, 5];
            c.Value = "Est. word count";

            c = ws.Cells[0, 6];
            c.Value = "Est. character count";

            c = ws.Cells[0, 7];
            c.Value = "Image count";

            c = ws.Cells[0, 8];
            c.Value = "Embedded";

            c = ws.Cells[0, 9];
            c.Value = "Password";

            c = ws.Cells[0, 10];
            c.Value = "Comment";



            int sor = 1;
            foreach (FileObject tf in fileList)
            {
                c = ws.Cells[sor, 0];
                c.Value = tf.fullPath == null ? "" : tf.fullPath;

                c = ws.Cells[sor, 1];
                c.Value = tf.path == null ? "" : tf.path;

                c = ws.Cells[sor, 2];
                c.Value = tf.filename == null ? "" : tf.filename;

                c = ws.Cells[sor, 3];
                c.Value = tf.fileType == null ? "" : tf.fileType;

                c = ws.Cells[sor, 4];
                c.Value = tf.pageCount == 0 ? "" : tf.pageCount.ToString();

                c = ws.Cells[sor, 5];
                c.Value = tf.wordCount == 0 ? "" : tf.wordCount.ToString();

                c = ws.Cells[sor, 6];
                c.Value = tf.characterCount == 0 ? "" : tf.characterCount.ToString();

                c = ws.Cells[sor, 7];
                c.Value = tf.imageCount == 0 ? "" : tf.imageCount.ToString();

                c = ws.Cells[sor, 8];
                c.Value = tf.embeddedDocsCount == 0 ? "" : tf.embeddedDocsCount.ToString();

                c = ws.Cells[sor, 9];
                c.Value = tf.hasPassword ? "The document is protected by a password." : "";

                c = ws.Cells[sor, 10];
                c.Value = tf.comment;

                sor++;
            }

            ws.Cells.HideColumn(0);
            for (int k = 1; k < 10; k++)
            {
                ws.AutoFitColumn(k);
            }
            wb.Save(newXLSPath);
            System.Diagnostics.Process.Start(newXLSPath);
        }


    }
}