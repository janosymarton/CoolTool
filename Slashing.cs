using Aspose.Words;
using Aspose.Words.Tables;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace CoolTool
{
    internal class Slashing
    {
        private string wordFile;
        Document doc;
        bool Success = true;
        string separator;

        public Slashing(string file, string separator)
        {
            this.wordFile = file;
            this.separator = separator;
            string extension = Path.GetExtension(file);
            if (extension != ".docx" && extension != ".doc")
            {
                Log.AddLog("File extension not valid.", true);
                return;
            }

            doc = new Document(wordFile);
            Process();
            doc.Save(wordFile);
            Program.mainWindow.updateProgress(100);

            if (Success)
            {
                Log.AddLog("Succesfully converted");
            }
            else
            {
                Log.AddLog("Problem while converting", true);
            }

        }

        private void Process()
        {
            try
            {
                int NoOfSections = doc.Sections.Count;
                int currentSection = 0;

                foreach (Section section in doc.Sections)
                {

                    Node startNode = GenerateStartNode(section);
                    Node node = startNode;
                    int NoOfParagraphs = section.Body.ChildNodes.Count;
                    int currentParagraph = 0;
                    DocumentBuilder builder = new DocumentBuilder(doc);
                    bool isEndOfFile = false;

                    while (!isEndOfFile)
                    {
                        if (node.NodeType == NodeType.Paragraph)
                        {

                            Paragraph para = (Paragraph)node;
                            isEndOfFile = para == section.Body.LastParagraph;

                            if (para.Runs.Count > 0)
                            {
                                
                                List<Font> runFonts = new List<Font>();
                                List<string> runTexts = new List<string>();
                                Font[] runFontsArr;
                                string[] runTextsArr;

                                foreach (Run paraRun in para.Runs)
                                {
                                    if (!Regex.IsMatch(paraRun.Text, @"^\s*FORM"))
                                    {
                                        runFonts.Add(paraRun.Font);
                                        runTexts.Add(paraRun.Text);
                                        paraRun.Font.Hidden = true;
                                    }
                                    else
                                    {
                                        builder.MoveTo(paraRun.PreviousSibling);
                                        runFontsArr = runFonts.ToArray();
                                        runTextsArr = runTexts.ToArray();

                                        writeSlashPart(ref builder, runFontsArr, runTextsArr);

                                        runFonts.Clear();
                                        runTexts.Clear();
                                    }

                                }

                                builder.MoveTo(node);
                                runFontsArr = runFonts.ToArray();
                                runTextsArr = runTexts.ToArray();

                                writeSlashPart(ref builder, runFontsArr, runTextsArr);
                                
                            }
                    
                        }
                        else if (node.NodeType == NodeType.Table)
                        {
                            Table table = (Table)node;

                            foreach (Row row in table.Rows)
                            {

                                foreach (Cell cell in row.Cells)
                                {
                                    Node tblStartNode = cell.FirstChild;
                                    for (Node tblNode = tblStartNode; tblNode != null; tblNode = tblNode.NextSibling)
                                    {
                                        if (tblNode.NodeType == NodeType.Paragraph)
                                        {
                                            builder.MoveTo(tblNode);
                                            Paragraph para = (Paragraph)tblNode;

                                            if (para.Runs.Count > 0)
                                            {

                                                List<Font> runFonts = new List<Font>();
                                                List<string> runTexts = new List<string>();
                                                Font[] runFontsArr;
                                                string[] runTextsArr;

                                                foreach (Run paraRun in para.Runs)
                                                {
                                                    if (!Regex.IsMatch(paraRun.Text, @"^\s*FORM"))
                                                    {
                                                        runFonts.Add(paraRun.Font);
                                                        runTexts.Add(paraRun.Text);
                                                        paraRun.Font.Hidden = true;
                                                    }
                                                    else
                                                    {
                                                        builder.MoveTo(paraRun.PreviousSibling);
                                                        runFontsArr = runFonts.ToArray();
                                                        runTextsArr = runTexts.ToArray();

                                                        writeSlashPart(ref builder, runFontsArr, runTextsArr);

                                                        runFonts.Clear();
                                                        runTexts.Clear();
                                                    }

                                                }

                                                builder.MoveTo(tblNode);
                                                runFontsArr = runFonts.ToArray();
                                                runTextsArr = runTexts.ToArray();

                                                writeSlashPart(ref builder, runFontsArr, runTextsArr);

                                            }
                                        }
                                    }
                                }
                            }
                        }
                        currentParagraph++;
                        double dblKeszultseg = ((70.0 / Convert.ToDouble(NoOfSections)) /
                            Convert.ToDouble(NoOfParagraphs)) * Convert.ToDouble(currentParagraph);
                        Program.mainWindow.updateProgress(30 + Convert.ToInt32(Math.Round(dblKeszultseg)));
                        node = node.NextSibling;

                    }
                    currentSection++;

                }
            }
            catch (Exception ex)
            {
                Log.AddLog("Tabulator position setting error.\r\n" + ex.Message, true);
                Success = false;
            }

        }

        private void writeSlashPart(ref DocumentBuilder builder, Font[] runFontsArr, string[] runTextsArr)
        {
            if(runTextsArr.Length != runFontsArr.Length)
            {
                builder.Write("ERROR");
                return;
            }

            string fullText = "";
            for (int i = 0; i < runTextsArr.Length; i++)
            {
                fullText += runTextsArr[i];
            }
            if (Regex.IsMatch(fullText, @"^\s*$"))
            {
                return;
            }

            builder.Font.Hidden = false;
            builder.Write(separator);

            runTextsArr[0] = runTextsArr[0].TrimStart();
            
            for (int i = 0; i < runFontsArr.Length; i++)
            {
                Font font = builder.Font;
                CopyFont(ref font, runFontsArr[i]);
                builder.Font.Hidden = false;
                builder.Write(runTextsArr[i]);
            }
        }

        private void CopyFont(ref Font dest, Font source)
        {
            if (source.GetType() != dest.GetType())

                throw new ArgumentException("All objects must be of the same type");



            // Iterate through each property in the source object.

            foreach (PropertyInfo prop in source.GetType().GetProperties())

            {

                // Skip indexed access items. Skip setting the internals of a style as these should not be changed.

                if (prop.Name == "Item" || prop.Name == "Style")

                    continue;



                object value;



                // Wrap this call as it can throw an exception. Skip if thrown

                try

                {

                    value = prop.GetValue(source, null);

                }



                catch (Exception)

                {

                    continue;

                }



                // Skip if value can not be retrieved.

                if (value != null)

                {

                    // If this property returns a class which belongs to the

                    if (prop.CanWrite)

                    {

                        prop.SetValue(dest, prop.GetValue(source, null), null);

                    }

                }

            }
        }

        private Node GenerateStartNode(Section section)
        {

            Node startNode = section.Body.FirstChild;
            while (startNode.NodeType != NodeType.Paragraph && startNode.NodeType != NodeType.Table)
            {
                startNode = section.Body.NextSibling;
            }

            if (startNode.ParentNode.NodeType != NodeType.Body)
            {
                throw new System.ApplicationException("Starting node must be a paragraph or a table in the main story of the document.");
            }

            return startNode;

        }

    }
}