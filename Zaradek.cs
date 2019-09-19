using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Layout;
using Aspose.Words.Tables;
using System.Reflection;

namespace CoolTool
{
    class Zaradek
    {
        private string wordFile;
        private string WMText;
        Document doc;
        bool Success = true;
        

        public Zaradek(string file, string WMText)
        {
            this.wordFile = file;
            this.WMText = WMText;
            
            string extension = Path.GetExtension(file);
            if (extension != ".docx" && extension != ".doc")
            {
                Log.AddLog("File extension not valid.", true);
                return;
            }

            RunConverterAspose();

            if (Success)
            {
                Log.AddLog("Succesfully converted");
            }
            else
            {
                Log.AddLog("Problem while converting", true);
            }
        }

        private void RunConverterAspose()
        {
            doc = new Document(wordFile);
            
            AddWatermarkAspose();
            doc.Save(wordFile);
            Program.mainWindow.updateProgress(25);
            
            AddTabulatorPosition();
            doc.Save(wordFile);
            Program.mainWindow.updateProgress(60);

            FillUpWithEmptyLines();
            doc.Save(wordFile);
            Program.mainWindow.updateProgress(100);

        }


        private void AddTabulatorPosition()
        {
            try
            {
                int NoOfSections = doc.Sections.Count;
                int currentSection = 0;

                foreach (Section section in doc.Sections)
                {
                    
                    Node startNode = GenerateStartNode(section);
                    Node node;
                    int NoOfParagraphs = section.Body.ChildNodes.Count;
                    int currentParagraph = 0;

                    for (node = startNode; node != null; node = node.NextSibling)
                    {
                        if (node.NodeType == NodeType.Paragraph)
                        {
                            DocumentBuilder builder = new DocumentBuilder(doc);
                            PageSetup pageSetup = builder.PageSetup;
                            builder.MoveTo(node);
                            Double tabPos = pageSetup.PageWidth - pageSetup.LeftMargin - pageSetup.RightMargin;
                            InsertTabs(node, tabPos, builder.ListFormat.IsListItem);
                        }
                        else if (node.NodeType == NodeType.Table)
                        {
                            Table table = (Table)node;
                            foreach (Row row in table.Rows)
                            {
                                foreach (Cell cell in row.Cells)
                                {
                                    Double tabPos =  cell.CellFormat.Width 
                                        - cell.CellFormat.RightPadding - cell.CellFormat.LeftPadding;
                                    Node tblStartNode = cell.FirstChild;
                                    
                                    for (Node tblNode = tblStartNode; tblNode != null; tblNode = tblNode.NextSibling)
                                    {
                                        if (tblNode.NodeType == NodeType.Paragraph)
                                        {
                                            DocumentBuilder builder = new DocumentBuilder(doc);
                                            builder.MoveTo(tblNode);
                                            InsertTabs(tblNode, tabPos, builder.ListFormat.IsListItem);
                                        }
                                    }
                                }
                            }
                        }
                        currentParagraph++;
                        double dblKeszultseg = ((20.0 / Convert.ToDouble( NoOfSections )) / 
                            Convert.ToDouble( NoOfParagraphs )) * Convert.ToDouble( currentParagraph );
                        Program.mainWindow.updateProgress(10 + Convert.ToInt32( Math.Round( dblKeszultseg )));
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

        private void InsertTabs(Node currentNode, double tabPos, bool islistitem)
        {
            Paragraph para = (Paragraph)currentNode;
            TabStop newtab = new TabStop(tabPos, Aspose.Words.TabAlignment.Left, TabLeader.Dashes);
            int noOfTabs = NoOfTabs(para) + (islistitem?1:0);
                
            for (int i = 0; i < para.ParagraphFormat.TabStops.Count; i++)
            {
                
                TabStop oldTab = para.ParagraphFormat.TabStops[i];
                if (oldTab.Position >= tabPos || noOfTabs - 1 < i)
                {
                    para.ParagraphFormat.TabStops.RemoveByIndex(i);
                }
            }
            para.ParagraphFormat.TabStops.Add(newtab);
        }

        private int NoOfTabs(Paragraph para)
        {
            string source = para.GetText();
            int count = 0;
            foreach (char c in source)
            {
                if (c == ControlChar.TabChar)
                {
                    count++;
                }
            }
            return count;
        }

        private void FillUpWithEmptyLines()
        {
            try
            {
                int NoOfSections = doc.Sections.Count;
                int currentSection = 0;

                foreach (Section section in doc.Sections)
                {
                    DocumentBuilder builder = new DocumentBuilder(doc);
                    Node startNode = GenerateStartNode(section);
                    Node node;
                    
                    int NoOfParagraphs = section.Body.ChildNodes.Count;
                    int currentParagraph = 0;

                    for (node = startNode; node != null; node = node.NextSibling)
                    {
                        
                        if (node.NodeType == NodeType.Paragraph)
                        {
                            builder.MoveTo(node);
                            Paragraph para = (Paragraph)node;
                            // string test = para.GetText();
                            Aspose.Words.Font font;
                            ParagraphFormat paragraphFormat;
                            InitFormats(para, ref builder, out font, out paragraphFormat);

                            Paragraph paraToInsert;
                            int paraRuns;
                            bool isEndOfFile;

                            if (para.GetText().IndexOf(ControlChar.PageBreakChar) >= 0)
                            {
                                LayoutCollector layoutCollector = new LayoutCollector(doc);
                                doc.UpdatePageLayout();
                                int origPage = layoutCollector.GetEndPageIndex(node);
                                int newpage = origPage;
                                
                                while (newpage == origPage)
                                {
                                    paraToInsert = (Paragraph)node;
                                    isEndOfFile = paraToInsert == section.Body.LastParagraph;
                                    paraRuns = paraToInsert.Runs.Count;
                                    if (paraRuns != 0 && !isEndOfFile)
                                    {
                                        builder.MoveTo(paraToInsert.Runs[paraRuns - 1]);
                                    }
                                    else
                                    {
                                        builder.MoveTo(paraToInsert);
                                    }

                                    builder.Writeln(ControlChar.Tab);
                                    
                                    if(builder.ListFormat.IsListItem)
                                    {
                                        builder.ListFormat.RemoveNumbers();
                                    }
                                    
                                    
                                    layoutCollector = new LayoutCollector(doc);
                                    doc.UpdatePageLayout();
                                    node = node.NextSibling;
                                    
                                    newpage = layoutCollector.GetEndPageIndex(node);
                                }

                                Node nodeToRemove = node.PreviousSibling;
                                Paragraph paraToRemove = (Paragraph)nodeToRemove;
                                string TextToRemove = paraToRemove.GetText();

                                if (TextToRemove == "\t\r")
                                {
                                    paraToInsert = (Paragraph)node;
                                    nodeToRemove.Remove();
                                    isEndOfFile = paraToInsert == section.Body.LastParagraph;
                                    paraRuns = paraToInsert.Runs.Count;
                                    if (paraRuns != 0 && !isEndOfFile)
                                    {
                                        builder.MoveTo(paraToInsert.Runs[paraRuns - 1]);
                                    }
                                    else if (paraRuns != 0 && isEndOfFile)
                                    {
                                        builder.MoveTo(paraToInsert.Runs[paraRuns]);
                                    }
                                    else
                                    {
                                        builder.MoveTo(paraToInsert);
                                    }
                                    builder.Write(ControlChar.Tab);
                                }
                                else
                                {
                                    node = node.PreviousSibling;
                                    node.NextSibling.Remove();
                                }
                                    
                            }
                            else
                            {
                                builder.MoveTo(node);
                                builder.Write(ControlChar.Tab);
                            }
                        }
                        else if (node.NodeType == NodeType.Table)
                        {
                            Table table = (Table)node;
                            LayoutCollector layoutCollector = new LayoutCollector(doc);
                            
                            foreach (Row row in table.Rows)
                            {
                               
                                foreach (Cell cell in row.Cells)
                                {
                                    Node tblStartNode = cell.FirstChild;
                                    for (Node tblNode = tblStartNode; tblNode != null; tblNode = tblNode.NextSibling)
                                    {
                                        builder.MoveTo(tblNode);
                                        if (tblNode.NodeType == NodeType.Paragraph)
                                        {
                                            Paragraph tblPara = (Paragraph)tblNode;
                                            Aspose.Words.Font font;
                                            ParagraphFormat paragraphFormat;
                                            InitFormats(tblPara, ref builder, out font, out paragraphFormat);
                                            builder.MoveTo(tblNode);
                                            builder.Write(ControlChar.Tab);

                                        }
                                    }

                                }
                            }
                        }
                        currentParagraph++;
                        double dblKeszultseg = ((70.0 / Convert.ToDouble( NoOfSections )) / 
                            Convert.ToDouble( NoOfParagraphs )) * Convert.ToDouble( currentParagraph );
                        Program.mainWindow.updateProgress(30 + Convert.ToInt32( Math.Round( dblKeszultseg )));
                    }
                    currentSection++;

                }
            }
            catch (Exception ex)
            {
                Log.AddLog("Adding empty lines error.\r\n" + ex.Message, true);
            }

        }

        private void InitFormats(Paragraph para, ref DocumentBuilder builder, out Aspose.Words.Font font, 
                                    out ParagraphFormat paragraphFormat)
        {
            Aspose.Words.Font origFont = para.ParagraphBreakFont;

            font = builder.Font;
            font.Size = origFont.Size;
            font.Bold = origFont.Bold;
            font.Color = origFont.Color;
            font.Name = origFont.Name;
            font.Underline = origFont.Underline;


            paragraphFormat = builder.ParagraphFormat;
            //paragraphFormat.FirstLineIndent = para.ParagraphFormat.FirstLineIndent;
            //paragraphFormat.Alignment = para.ParagraphFormat.Alignment;
            //paragraphFormat.KeepTogether = para.ParagraphFormat.KeepTogether;
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

        private void AddWatermarkAspose()
        {
            try
            {
                Paragraph watermarkPara = new Paragraph(doc);
                
                Shape watermark = new Shape(doc, ShapeType.TextPlainText);

                watermark.TextPath.Text = WMText;
                watermark.TextPath.FontFamily = "Times New Roman";
                watermark.Width = 500;
                watermark.Height = 100;
                watermark.Fill.Color = Color.LightGray;  // Try LightGray to get more Word-style watermark
                watermark.StrokeColor = Color.LightGray; // Try LightGray to get more Word-style watermark
                watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
                watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page;
                watermark.WrapType = WrapType.None;
                watermark.VerticalAlignment = VerticalAlignment.Center;
                watermark.HorizontalAlignment = Aspose.Words.Drawing.HorizontalAlignment.Center;
                watermark.Rotation = -40;

                watermarkPara.AppendChild(watermark);
                
                Shape Logo = new Shape(doc, ShapeType.Image);

                Assembly assembly = Assembly.GetExecutingAssembly();
                Stream imageStream = assembly.GetManifestResourceStream("CoolTool.ZaradekLogo.jpg");
                Logo.ImageData.SetImage(imageStream);
                Logo.Width = ConvertUtil.MillimeterToPoint(25);
                Logo.Height = ConvertUtil.MillimeterToPoint(25);
                Logo.RelativeHorizontalPosition = RelativeHorizontalPosition.Margin;
                Logo.RelativeVerticalPosition = RelativeVerticalPosition.Line;
                Logo.VerticalAlignment = Aspose.Words.Drawing.VerticalAlignment.Top;
                Logo.HorizontalAlignment = Aspose.Words.Drawing.HorizontalAlignment.Right;
                

                Logo.WrapType = WrapType.None;
                
                watermarkPara.AppendChild(Logo);
                
                foreach (Section sect in doc.Sections)
                {
                    sect.PageSetup.TopMargin = ConvertUtil.MillimeterToPoint(37.5);
                    sect.PageSetup.HeaderDistance = ConvertUtil.MillimeterToPoint(8);
                    InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HeaderPrimary);
                    InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HeaderFirst);
                    InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HeaderEven);
                }
            }

            catch (Exception ex)
            {
                Log.AddLog("Watermark and logo error.\r\n" + ex.Message, true);
            }

        }


        private void InsertWatermarkIntoHeader(Paragraph watermarkPara, Section sect, HeaderFooterType headerFooterType)
        {
            HeaderFooter header = sect.HeadersFooters[headerFooterType];

            if (header == null)
            {
                header = new HeaderFooter(sect.Document, headerFooterType);
                sect.HeadersFooters.Add(header);
            }

            Node newNode = header.AppendChild(watermarkPara.Clone(true));
        }



    }
}
