using Aspose.Pdf;
using System.IO;
using System;
using Aspose.Pdf.Annotations;
using Aspose.Words;
using System.Reflection;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace CoolTool
{
    internal class ZaradekFinisher
    {
        private string basePath;
        private Dictionary<int, ZaradekFinishDoc> mergableFiles = new Dictionary<int, ZaradekFinishDoc>();
        private string outputFile;
        private string tempFile;
        private int stampMode;

        public ZaradekFinisher(string basePath, int stampMode)
        {
            this.basePath = basePath;
            this.stampMode = stampMode;
            outputFile = Path.Combine(basePath, "mergedFile.pdf");

            int i = 0;
            while (File.Exists(outputFile))
            {
                i++;
                outputFile = Path.Combine(basePath, "mergedFile_" + i + ".pdf");
            }
            tempFile = Path.Combine(basePath, "tempFile.pdf");
            if (File.Exists(tempFile))
                File.Delete(tempFile);

            Process();

            File.Delete(tempFile);
        }

        private void Process()
        {
            FillFiles();
            Program.mainWindow.updateProgress(25);

            Log.AddLog("Fájlok azonosítása kész. DOC(X) fájlok konvertálása PDF-é folyamatban.");

            foreach (int filePosition in mergableFiles.Keys)
            {
                if (String.IsNullOrEmpty(mergableFiles[filePosition].filePDF))
                {
                    if (!ConvertDOC2PDF(filePosition))
                    {
                        Log.AddLog("Hiba a PDF-é konvertálás közben: " + mergableFiles[filePosition].fileDocx +
                                            " - Feldolgozás leáll.", true);
                        return;
                    }
                    else
                        Log.AddLog("A célnyelvi fájl kovertálása PDF-é sikeres: " + mergableFiles[filePosition].filePDF);
                }
            }

            Program.mainWindow.updateProgress(50);

            Log.AddLog("A fájlok összeillesztése folyamatban...");

            if (MergeFiles())
            {
                Program.mainWindow.updateProgress(75);
                Log.AddLog("Fájlok összefűzése befejeződött.");
            }
            else
            {
                return;
            }

            if (stampMode != 3)
            {
                Log.AddLog("Pecsét beillesztése folyamatban...");
                InsertStamp();
            }
            else
            {
                File.Move(tempFile, outputFile);
            }

            Program.mainWindow.updateProgress(100);
            Log.AddLog("A mappa feldolgozása befejeződött: " + outputFile);
        }

        private void FillFiles()
        {
            string[] filesInFolder = Directory.GetFiles(basePath);
            foreach (string fileInFolder in filesInFolder)
            {
                try
                {
                    if (Regex.IsMatch(Path.GetFileNameWithoutExtension(fileInFolder).Substring(0, 3), @"\d{2}_") &&
                        (Path.GetExtension(fileInFolder).ToLower() == ".pdf" || Path.GetExtension(fileInFolder).ToLower() == ".docx"
                        || Path.GetExtension(fileInFolder).ToLower() == ".doc"))
                    {
                        int position;
                        if (Int32.TryParse(Path.GetFileNameWithoutExtension(fileInFolder).Substring(0, 2).TrimStart('0'), out position))
                        {
                            if (!mergableFiles.ContainsKey(position))
                            {
                                mergableFiles.Add(position, new ZaradekFinishDoc(fileInFolder));
                            }
                            else if (String.IsNullOrEmpty( mergableFiles[position].filePDF) &&
                                Path.GetFileNameWithoutExtension(mergableFiles[position].fileDocx) == Path.GetFileNameWithoutExtension(fileInFolder))
                            {
                                mergableFiles[position].filePDF = fileInFolder;
                            }
                            else
                            {
                                Log.AddLog("Azonos sorszámú fájlok: " + fileInFolder + " - " + mergableFiles[position].fullPath + 
                                    ". Az előbbi fájl figyelmen kívül hagyva.", true);
                            }
                        }
                        else
                        {
                            Log.AddLog("Sorszám nem értelmezhető: " + fileInFolder, true);
                        }
                    }
                    else
                    {
                        Log.AddLog("Nem megfelelő fájlnév a mappában: " + fileInFolder, true);
                    }
                }
                catch (Exception ex)
                {
                    Log.AddLog("Hiba a fájl azonosítása során: " + fileInFolder + " - " + ex.Message, true);
                }
            }
        }

        private bool ConvertDOC2PDF(int filePosition)
        {
            string sourceFile = mergableFiles[filePosition].fileDocx;
            try
            {
                Aspose.Words.Document doc = new Aspose.Words.Document(sourceFile);
                string convertedPDF = Path.Combine(basePath, Path.GetFileNameWithoutExtension(sourceFile) + ".pdf");
                doc.Save(convertedPDF);
                mergableFiles[filePosition].filePDF = convertedPDF;
            }
            catch (Exception ex)
            {
                Log.AddLog("DOC(X) fájl megnyitása sikertelen: " + sourceFile + " - " + ex.Message, true);
                return false;
            }
            return true;
        }


        private bool MergeFiles()
        {
            try
            {
                SortedDictionary<int, ZaradekFinishDoc> mergingFiles = new SortedDictionary<int, ZaradekFinishDoc>(mergableFiles);

                Aspose.Pdf.Document masterDoc = new Aspose.Pdf.Document();
                bool isFirst = true;

                foreach (int filePosition in mergingFiles.Keys)
                {
                    if (isFirst)
                    {
                        isFirst = false;
                        masterDoc = new Aspose.Pdf.Document(mergingFiles[filePosition].filePDF);
                    }
                    else
                    {
                        Aspose.Pdf.Document addPDFDoc = new Aspose.Pdf.Document(mergingFiles[filePosition].filePDF);
                        masterDoc.Pages.Add(addPDFDoc.Pages);
                    }
                }

                masterDoc.Save(tempFile);
            }
            catch (Exception ex)
            {
                Log.AddLog("A PDF fájlok összefűzése közben hiba történt - " + ex.Message, true);
                return false;
            }
            return true;
        }

        private void InsertStamp()
        {
            try
            {
                double imageSize = 115;
                double xShift = 10;
                double yShift = 10;


                Rectangle imageRectangle = new Rectangle(xShift, yShift, xShift + imageSize, yShift + imageSize);
                using (Aspose.Pdf.Document document = new Aspose.Pdf.Document(tempFile))
                {
                    Assembly assembly = Assembly.GetExecutingAssembly();

                    using (var imageStream = assembly.GetManifestResourceStream("CoolTool.pecsét.png"))
                    {
                        XImage image = null;
                        foreach (Page page in document.Pages)
                        {
                            double xCoor;
                            switch (stampMode)
                            {
                                case 1:
                                    xCoor = page.PageInfo.Width - xShift - imageSize;
                                    break;
                                case 2:
                                    xCoor = xShift;
                                    break;
                                default:
                                    xCoor = page.PageInfo.Width - xShift - imageSize;
                                    break;
                            }

                            WatermarkAnnotation annotation = new WatermarkAnnotation(page, page.Rect);
                            XForm form = annotation.Appearance["N"];
                            form.BBox = page.Rect;

                            string name;
                            if (image == null)
                            {
                                name = form.Resources.Images.Add(imageStream);
                                image = form.Resources.Images[name];
                            }
                            else
                            {
                                name = form.Resources.Images.Add(image);
                            }

                            form.Contents.Add(new Operator.GSave());
                            form.Contents.Add(new Operator.ConcatenateMatrix(new Matrix(imageSize, 0, 0, imageSize, xCoor, yShift)));
                            form.Contents.Add(new Operator.Do(name));
                            form.Contents.Add(new Operator.GRestore());

                            page.Annotations.Add(annotation, false);
                        }
                    }

                    document.Save(outputFile);
                }
            }
            catch (Exception ex)
            {
                Log.AddLog("Hiba a pecsét beillesztése során - " + ex.Message, true);
            }
        }
    }

    internal class ZaradekFinishDoc
    {
        public string filePDF;
        public string fileDocx;
        public string fullPath;

        public ZaradekFinishDoc(string fullPath)
        {
            this.fullPath = fullPath;
            if (Path.GetExtension(fullPath) == ".pdf")
            {
                filePDF = fullPath;
            }
            else if (Path.GetExtension(fullPath) == ".docx" || Path.GetExtension(fullPath) == ".doc")
            {
                fileDocx = fullPath;
            }
            else
                throw new Exception("Nem megfelelő kiterjesztés.");
        }
    }
}