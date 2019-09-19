
using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Text;

namespace CoolTool
{
    internal class SignatureGenerator
    {
        private string fullXLSPath;
        private string outputPath;
        private string tempPath;
        private string specialMessageEn;
        private string specialMessageHu;

        public SignatureGenerator(string filePath, string specialMessageEn, string specialMessageHu)
        {
            try
            {
                this.fullXLSPath = filePath;
                this.specialMessageEn = specialMessageEn;
                this.specialMessageHu = specialMessageHu;

                int i = 0;
                do
                {
                    outputPath = Path.Combine(Path.GetDirectoryName(fullXLSPath),
                                                "Signatures" + (i > 0 ? ("_" + i.ToString()) : ""));
                    i++;
                } while (Directory.Exists(outputPath));

                Directory.CreateDirectory(outputPath);

                tempPath = Path.Combine(outputPath, "_tmp");
                Directory.CreateDirectory(tempPath);

                string zipFilePath = Path.Combine(tempPath, "signaturetemplates.zip");
                string resourceName = "CoolTool.signaturetemplates.zip";
                using (var resource = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
                {

                    using (var zipFile = new FileStream(zipFilePath, FileMode.Create, FileAccess.Write))
                    {
                        resource.CopyTo(zipFile);
                    }
                }

                ZipFile.ExtractToDirectory(zipFilePath, tempPath);
                File.Delete(zipFilePath);
                Directory.Move(Path.Combine(tempPath, "edimart_images"), Path.Combine(outputPath, "edimart_images"));

                ProcessSignatureFile();

                Directory.Delete(tempPath, true);
                Directory.Delete(Path.Combine(outputPath, "edimart_images"), true);
                Program.mainWindow.updateProgress(100);
            }
            catch (Exception ex)
            {
                Log.AddLog("Signature processing problem: " + ex.Message, true);
            }

        }

        private void ProcessSignatureFile()
        {
            List<UserName> userList = new List<UserName>();
            try
            {
                Workbook wb = new Workbook(fullXLSPath);
                WorksheetCollection wsc = wb.Worksheets;
                Worksheet ws = wb.Worksheets[0];
                int lastRow = ws.Cells.GetLastDataRow(0);

                for (int i = 1; i <= lastRow; i++)
                {
                    string nameHu = ws.Cells[i, 0].Value != null ? ws.Cells[i, 0].Value.ToString() : "";
                    string nameEn = ws.Cells[i, 1].Value != null ? ws.Cells[i, 1].Value.ToString() : "";
                    string positionHu = ws.Cells[i, 2].Value != null ? ws.Cells[i, 2].Value.ToString() : "";
                    string positionEn = ws.Cells[i, 3].Value != null ? ws.Cells[i, 3].Value.ToString() : "";
                    string phoneNumber = ws.Cells[i, 4].Value != null ? ws.Cells[i, 4].Value.ToString() : "";
                    string skypeAccount = ws.Cells[i, 5].Value != null ? ws.Cells[i, 5].Value.ToString() : "";
                    string emailAddress = ws.Cells[i, 6].Value != null ? ws.Cells[i, 6].Value.ToString() : "";

                    userList.Add(new UserName(nameHu, nameEn, positionHu, positionEn, phoneNumber, skypeAccount, emailAddress));
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Problem while reading data XLS: " + ex.Message);
            }

            foreach (UserName userRecord in userList)
            {
                CreateSignatureFile(userRecord, "signature_template_EN.htm", "EN", Encoding.GetEncoding("windows-1250"));
                CreateSignatureFile(userRecord, "signature_template_HU.htm", "HU", Encoding.GetEncoding("windows-1250"));
                CreateSignatureFile(userRecord, "signature_template_EN.txt", "EN", Encoding.GetEncoding("iso-8859-2"));
                CreateSignatureFile(userRecord, "signature_template_HU.txt", "HU", Encoding.GetEncoding("iso-8859-2"));
                CreateImageFolder(userRecord, "EN");
                CreateImageFolder(userRecord, "HU");
            }

            Log.AddLog("Signature files generated " + outputPath);

        }

        private void CreateImageFolder(UserName userRecord, string lang)
        {
            string newFolderPath = userRecord.nameHu + " " + lang + "_elemei";
            string imageFolder = Path.Combine(outputPath, "edimart_images");
            Directory.CreateDirectory(Path.Combine(outputPath, newFolderPath));
            foreach (string file in Directory.GetFiles(imageFolder))
            {
                string imageFilePath = Path.Combine(outputPath, newFolderPath, Path.GetFileName(file));
                if (!File.Exists(imageFilePath))
                {
                    File.Copy(file, imageFilePath);
                }
            }
        }

        private void CreateSignatureFile(UserName userRecord, string templateName, string lang, Encoding enc)
        {

            string templatePath = Path.Combine(tempPath, templateName);
            string extension = Path.GetExtension(templatePath);
            string outputFilePath = Path.Combine(outputPath, userRecord.nameHu + " " + lang + extension);
            string folderHu = userRecord.nameHu.Replace(" ", "%20") + "%20HU_elemei";
            string folderEn = userRecord.nameHu.Replace(" ", "%20") + "%20EN_elemei";

            try
            {
                using (StreamReader input = new StreamReader(templatePath, enc))
                {
                    using (StreamWriter output = new StreamWriter(outputFilePath, false, enc))
                    {
                        string content;
                        output.AutoFlush = true;
                        content = input.ReadToEnd();
                        content = content.Replace("#namehu#", userRecord.nameHu);
                        content = content.Replace("#nameen#", userRecord.nameEn);
                        content = content.Replace("#positionen#", userRecord.positionEn);
                        content = content.Replace("#positionhu#", userRecord.positionHu);
                        content = content.Replace("#mobile#", userRecord.phoneNumber);
                        content = content.Replace("#skype#", userRecord.skypeAccount);
                        content = content.Replace("#email#", userRecord.emailAddress);
                        content = content.Replace("#specialmessageen#", specialMessageEn);
                        content = content.Replace("#specialmessagehu#", specialMessageHu);
                        content = content.Replace("#folderhu#", folderHu);
                        content = content.Replace("#folderen#", folderEn);

                        output.Write(content.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                Log.AddLog("Signature file creation error - " + outputFilePath + ": " + ex.Message, true);
            }

        }
    }

    class UserName
    {
        public string emailAddress;
        public string nameEn;
        public string nameHu;
        public string phoneNumber;
        public string positionEn;
        public string positionHu;
        public string skypeAccount;

        public UserName(string nameHu, string nameEn, string positionHu, string positionEn, string phoneNumber, string skypeAccount, string emailAddress)
        {
            this.nameHu = nameHu;
            this.nameEn = nameEn;
            this.positionHu = positionHu;
            this.positionEn = positionEn;
            this.phoneNumber = phoneNumber;
            this.skypeAccount = skypeAccount;
            this.emailAddress = emailAddress;
        }
    }
}