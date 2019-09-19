using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Data;
using System.ServiceModel;

namespace CoolTool
{
    class LangCopier
    {
        private string sInputPath;
        private Language[] languages;


        public LangCopier(string file, System.Windows.Forms.CheckedListBox.CheckedItemCollection checkedItemCollection)
        {
            Program.mainWindow.updateProgress(0);
            this.sInputPath = file;
            this.languages = LanguageHelper.GetLanguageList(checkedItemCollection);
            
            if (languages.Length == 0)
            {
                Log.AddLog("No target language was selected.", true);
                return;
            }
            LangCopyProcessor();
        }

        private void LangCopyProcessor()
        {
            string strOutputFile = "";

            int iFile = 0;
            int iConv = 0;

            if (sInputPath.Substring(sInputPath.Length) != "\\")
                sInputPath += "\\";

            string[] filePaths = Directory.GetFiles(sInputPath, "*.mqxlz");
            string strTempFolder = Path.GetDirectoryName(sInputPath) + @"\tmp";

            int langCount = languages.Length;
            int langCounter = 0;
            int fileCount = filePaths.Length;
            int fileCounter = 0;
            double progress = 0;
                        
            foreach (Language newTargetLang in languages)
            {
                langCounter++;
                string stroutputfolder = Path.GetDirectoryName(sInputPath) + @"\TargetFiles\" + newTargetLang.ISOCode;
                
                if(!Directory.Exists(stroutputfolder))
                    Directory.CreateDirectory(stroutputfolder);

                foreach (string strInputFile in filePaths)
                {
                    fileCounter++;
                    if(Directory.Exists(strTempFolder))
                    {
                        Directory.Delete(strTempFolder);
                    }
                    Directory.CreateDirectory(strTempFolder);
                    string exportXLZfile = stroutputfolder + @"\" + Path.GetFileName(strInputFile);
                    
                    ZipFile.ExtractToDirectory(strInputFile, strTempFolder);
                    string[] xlfPaths = Directory.GetFiles(strTempFolder, "*.mqxliff");
                    int xliffCount = xlfPaths.Length;
                    int xliffCounter = 0;

                    foreach (string xlf in xlfPaths)
                    {
                        xliffCounter++;
                        strOutputFile = Path.GetDirectoryName(xlf) + "tmp_" + Path.GetFileName(xlf);
                        double increment = 100 / (Convert.ToDouble(fileCount * langCount * xliffCount));
                        iConv = SR(xlf, strOutputFile, newTargetLang.ISOCode, progress, increment);

                        if (iConv == 0)
                        {
                            File.Delete(xlf);
                            File.Move(strOutputFile, xlf);
                            ZipFile.CreateFromDirectory(strTempFolder, exportXLZfile);
                            Directory.Delete(strTempFolder, true);
                            Log.AddLog("File converted: " + strInputFile, false);
                            iFile++;
                        }
                        else
                        {
                            if (File.Exists(strOutputFile))
                            {
                                File.Delete(strOutputFile);
                            }
                            if(Directory.Exists(strTempFolder))
                            {
                                Directory.Delete(strTempFolder);
                            }
                            
                            Log.AddLog("File conversion failed: " + strInputFile, true);
                        }

                        progress = langCounter * (100 / langCount) + fileCounter * (100 / (fileCount * langCount)) + 
                                        xliffCounter * (100 / (fileCount * langCount * xliffCount));
                        Program.mainWindow.updateProgress(Convert.ToInt32(progress));
                        
                    }
                }

                    Log.AddLog("Number of files converted: " + iFile, false);
                }
            }

        
            private int SR(string strInputFile, string strOutputFile, string newTargetLang, double progress, double increment)
            {

            
            if (File.Exists(strOutputFile))
            {
                Console.WriteLine("The output file already exist. Please specify a new one: " + strOutputFile);
                return 1;
            }
           

            if (!File.Exists(strInputFile))
            {
                Console.WriteLine("The input file doesn't exist: " + strInputFile);
                return 2;
            }
            
            
            using (StreamReader input = new StreamReader(strInputFile))
            using (StreamWriter output = new StreamWriter(strOutputFile))
            {
                string content;
                output.AutoFlush = true;
                content = input.ReadToEnd();


                string pattern1 = @"(<file\s[^>]+?target-language="")(.+?)(""[^>]+?>)";
                content = replacestring(true, pattern1, @"$1" + newTargetLang + "$3", content);

                string pattern2 = @"(?i)(mq:(id|segmentguid)="")([A-F0-9]{8}(?:-[A-F0-9]{4}){3}-[A-F0-9]{12})("")";
                
                MatchCollection matches = Regex.Matches(content, pattern2);

                int matchCount = matches.Count;
                int matchCounter = 0;

                foreach (Match match in matches)
                {
                    matchCounter++;

                    string strOldGuid = match.Groups[3].Value;
                    string strStart = match.Groups[1].Value;
                    string strEnd = @"""";
                    
                    string strOldText = strStart + strOldGuid + strEnd;
                    string strNewText = strStart + Guid.NewGuid() + strEnd;
                    
                    content = replacestring(false, strOldText, strNewText, content);

                    Program.mainWindow.updateProgress(Convert.ToInt32( progress + matchCounter * increment / matchCount));
                }
                output.Write(content.ToString());
            }


            return 0;
        }

        private static string replacestring(bool isRegex, string strSearchExp, string strReplaceExp, string content)
        {
            string strResult = "";
            if (isRegex)
            {
                // regex based replace
                Regex rgx = new Regex(strSearchExp);
                return strResult = rgx.Replace(content, strReplaceExp);

            }
            else
            {
                // text based replace
                return strResult = content.Replace(strSearchExp, strReplaceExp);
            }



        }


    }
}
