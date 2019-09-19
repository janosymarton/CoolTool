using System;
using Aspose.Pdf;
using Aspose.Words;
using System.IO;
using Aspose.Pdf.Devices;
using Aspose.Pdf.Facades;

namespace CoolTool
{
    class PDFHelper
    {
        private string sPath;
        
        public PDFHelper(string file)
        {
            this.sPath = file;
            
            DoSplit();
        }


        internal void DoSplit()
        {
            
            string[] files = Directory.GetFiles(sPath, "*.pdf", SearchOption.AllDirectories);
            int fileCount = files.Length;
            int fileCounter = 0;

            foreach(string sourceFile in files)
            {
                Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(sourceFile);
                //Create a PdfConverter object
                PdfConverter converter = new PdfConverter();
                //Bind the input PDF file
                converter.BindPdf(sourceFile);
                // Specify the start page to be processed
                converter.StartPage = 1;
                // Specify the end page for processing
                converter.EndPage = pdfDocument.Pages.Count;
                // Create a Resolution object to specify the resolution of resultant image
                converter.Resolution = new Aspose.Pdf.Devices.Resolution(300);
                //Initialize the convertion process
                converter.DoConvert();

                string docPath = Path.Combine(Path.GetDirectoryName(sourceFile),
                        Path.GetFileNameWithoutExtension(sourceFile) + ".docx");
                Aspose.Words.Document doc = new Aspose.Words.Document();
                DocumentBuilder builder = new DocumentBuilder(doc);

                int pageCount = 0;
                //Check if pages exist and then convert to image one by one
                while (converter.HasNextImage())
                {
                    pageCount++;
                    string filePath = Path.Combine(Path.GetDirectoryName(sourceFile),
                        Path.GetFileNameWithoutExtension(sourceFile) + "_" + pageCount.ToString() + ".png");
                    using (MemoryStream imageStream = new MemoryStream())
                    {
                        converter.GetNextImage(imageStream, System.Drawing.Imaging.ImageFormat.Png);
                        imageStream.Position = 0;
                        System.IO.File.WriteAllBytes(filePath, imageStream.ToArray());
                    }
                    builder.InsertImage(filePath);
                    
                }
                // Close the PdfConverter instance and release the resources
                doc.Save(docPath);
                converter.Close();
                // Close the stream holding the image object
                fileCounter++;
                Program.mainWindow.updateProgress(Convert.ToInt32( Math.Floor( Convert.ToDouble(fileCounter) * (100 / Convert.ToDouble(fileCount)))));   
            }

            Program.mainWindow.updateProgress(100);
        }


        //internal static void DoProcess()
        //{
        //    DocSaveOptions saveOptions = new DocSaveOptions();
        //    saveOptions.RecognizeBullets = true;
        //    saveOptions.Mode = DocSaveOptions.RecognitionMode.Flow;
        //    saveOptions.AddReturnToLineEnd = false;

        //    using (Document doc = new Document(@"d:\temp\gyerekek\zsuzsi\kötelező olv\1. Richelle Mead - Vámpírakadémia.pdf"))
        //    {
        //        doc.Save(@"d:\temp\gyerekek\zsuzsi\kötelező olv\1. Richelle Mead - Vámpírakadémia.docx", saveOptions);
        //    }

        //}

    }
}