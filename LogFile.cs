using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Aspose.Words;
using Aspose.Cells;
using System.Drawing;
using System.Xml;
using System.Text.RegularExpressions;


namespace CoolTool
{
    class LogFile
    {
        private string logfile;

        bool Success = true;
        private char Separator;
        private FileTypes fileType;
        private Dictionary<string, TranslationFile> translationFiles = new Dictionary<string, TranslationFile>();
        private int[] grid;
        private Worksheet ws;
        private int TotalColumn;
        public Color[] columnColors;
        private Color c1 = Color.MistyRose;
        private Color c2 = Color.CadetBlue;
        private Color c3 = Color.LightCyan;
        private Color c4 = Color.Thistle;
        private Color c5 = Color.MediumTurquoise;
        private Color c6 = Color.LavenderBlush;
        private Color c7 = Color.Aquamarine;


        public LogFile(string csvfile, int gridIndex)
        {
            this.logfile = csvfile;
            
            this.grid = Program.grids.ElementAt(gridIndex).Value;
            string extension = Path.GetExtension(csvfile);
            if (extension.ToLower() != ".csv" && extension.ToLower() != ".html" && extension.ToLower() != ".xml")
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
            try
            {
                ParseHeader();
                Program.mainWindow.updateProgress(10);
                ParseData();
                Program.mainWindow.updateProgress(25);
                CreateXLS();

            }
            catch (Exception ex)
            {

                Log.AddLog("Converter error: " + ex.Message, true);
                Success = false;
            }
        }



        private void ParseHeader()
        {
            using (StreamReader reader = new StreamReader(File.OpenRead(logfile)))
            {
                string headerline = reader.ReadLine();
                if (headerline.StartsWith(@"<task"))
                {
                    fileType = FileTypes.StudioXML;
                }
                else if (headerline.StartsWith("<html"))
                {
                    if (headerline.Contains(@"xmlns:weightedCountsHelper=""WeightedCountsHelper""") || headerline.Contains("<html>"))
                        fileType = FileTypes.memoQHTML;
                    else
                    {
                        reader.ReadLine();
                        if (reader.ReadLine().Contains("Trados"))
                            fileType = FileTypes.StudioHTML;
                        else
                        {
                            Success = false;
                            throw new Exception("Invalid logfile format.");
                        }
                    }

                }
                else if (headerline.Contains("Statistics"))
                {
                    fileType = FileTypes.memoQ;
                    string dataline;
                    while (!(dataline = reader.ReadLine()).Contains("101%"))
                    {
                        Helper.FindSeparator(dataline, out Separator, out Success);
                    }
                }
                else if (Helper.FindSeparator(headerline, out Separator, out Success))
                {
                    string[] headerParts = headerline.Split(Separator);
                    if (headerParts[2] == "X-translated")
                        fileType = FileTypes.memoQAllInfo;
                    else if (headerParts[3] == "Context TM")
                        fileType = FileTypes.memoQTrados;
                    else
                    {
                        Success = false;
                        throw new Exception("Invalid logfile format.");
                    }
                }
                else
                {
                    Success = false;
                    throw new Exception("Invalid logfile format.");
                }
            }
        }

        private void ParseData()
        {
            switch (fileType)
            {
                case FileTypes.memoQAllInfo:
                case FileTypes.memoQTrados:
                    ParseCSV();
                    break;
                case FileTypes.memoQHTML:
                case FileTypes.StudioXML:
                case FileTypes.StudioHTML:
                    ParseHTML();
                    break;
                case FileTypes.memoQ:
                    ParsememoQText();
                    break;
                default:
                    Success = false;
                    throw new Exception("Invalid logfile format.");
            }

            if (translationFiles.Count > 0)
            {
                Helper.FindCommonPath(ref translationFiles);
            }
            else
            {
                Success = false;
                throw new Exception("Invalid logfile format or no file included.");
            }
        }

        private void ParseCSV()
        {
            using (StreamReader reader = new StreamReader(File.OpenRead(logfile)))
            {
                reader.ReadLine();
                reader.ReadLine();
                string dataline;
                while ((dataline = reader.ReadLine()) != null)
                {
                    string[] dataArray = dataline.Split(Separator);
                    if (!translationFiles.ContainsKey(dataArray[0]))
                    {
                        int[] wCounts = new int[9];
                        int total = 0;
                        switch (fileType)
                        {
                            case FileTypes.memoQAllInfo:
                                wCounts[0] = Convert.ToInt32(dataArray[3]);
                                wCounts[1] = Convert.ToInt32(dataArray[11]);
                                wCounts[2] = Convert.ToInt32(dataArray[19]);
                                wCounts[3] = Convert.ToInt32(dataArray[27]);
                                wCounts[4] = Convert.ToInt32(dataArray[35]);
                                wCounts[5] = Convert.ToInt32(dataArray[43]);
                                wCounts[6] = Convert.ToInt32(dataArray[51]);
                                wCounts[7] = Convert.ToInt32(dataArray[59]);
                                wCounts[8] = Convert.ToInt32(dataArray[67]);
                                total = Convert.ToInt32(dataArray[83]);
                                break;
                            case FileTypes.memoQTrados:
                                wCounts[0] = 0;
                                wCounts[1] = Convert.ToInt32(dataArray[4]);
                                wCounts[2] = Convert.ToInt32(dataArray[8]);
                                wCounts[3] = Convert.ToInt32(dataArray[12]);
                                wCounts[4] = Convert.ToInt32(dataArray[16]);
                                wCounts[5] = Convert.ToInt32(dataArray[20]);
                                wCounts[6] = Convert.ToInt32(dataArray[24]);
                                wCounts[7] = Convert.ToInt32(dataArray[28]);
                                wCounts[8] = Convert.ToInt32(dataArray[32]);
                                total = Convert.ToInt32(dataArray[36]);
                                break;
                            default:
                                throw new InvalidDataException("Unhandled filetype.");
                        }
                        TranslationFile tf = new TranslationFile(dataArray[0], wCounts, total, grid);
                        translationFiles.Add(dataArray[0], tf);
                    }
                    else
                    {
                        Log.AddLog("Duplicate file in logfile, file skipped: " + dataArray[0], true);
                    }

                }

            }
        }

        private void ParseHTML()
        {
            using (StreamReader reader = new StreamReader(File.OpenRead(logfile)))
            {
                string fullFile = reader.ReadToEnd();
                string rgx = "";
                switch (fileType)
                {
                    case FileTypes.memoQHTML:
                        rgx = @"<h2>Analysis</h2>\s*<table[^>]*>\s*<tr[^>]*>\s*<td[^>]*>Scope</td>\s*<td[^>]*>""" +
                            @"(?<filename>[^""]*)""\s*(\(""(?<fullpath>[^""]*)""\)(\s*\(\d+\))?)?</td>\s*</tr>.+?</table>\s*<table[^>]*><tr[^>]*>.+?</tr>\s*" +
                            @"<tr[^>]*>\s*<td[^>]*>[^<]*</td>\s*<td>\d+</td>\s*<td>(?<Total>\d+)</td>\s*<td>\d+</td>\s*<td>\d+</td>\s*<td>[\d,.]+</td>\s*</tr>\s*" +
                            @"<tr[^>]*>\s*<td[^>]*>[^<]*</td>\s*<td>\d+</td>\s*<td>(?<XTrans>\d+)</td>\s*<td>\d+</td>\s*<td>\d+</td>\s*<td>[\d,.]+</td>\s*</tr>\s*" +
                            @"<tr[^>]*>\s*<td[^>]*>[^<]*</td>\s*<td>\d+</td>\s*<td>(?<Reps>\d+)</td>\s*<td>\d+</td>\s*<td>\d+</td>\s*<td>[\d,.]+</td>\s*</tr>\s*" +
                            @"<tr[^>]*>\s*<td[^>]*>[^<]*</td>\s*<td>\d+</td>\s*<td>(?<CM>\d+)</td>\s*<td>\d+</td>\s*<td>\d+</td>\s*<td>[\d,.]+</td>\s*</tr>\s*" +
                            @"<tr[^>]*>\s*<td[^>]*>[^<]*</td>\s*<td>\d+</td>\s*<td>(?<pc100>\d+)</td>\s*<td>\d+</td>\s*<td>\d+</td>\s*<td>[\d,.]+</td>\s*</tr>\s*" +
                            @"<tr[^>]*>\s*<td[^>]*>[^<]*</td>\s*<td>\d+</td>\s*<td>(?<pc95>\d+)</td>\s*<td>\d+</td>\s*<td>\d+</td>\s*<td>[\d,.]+</td>\s*</tr>\s*" +
                            @"<tr[^>]*>\s*<td[^>]*>[^<]*</td>\s*<td>\d+</td>\s*<td>(?<pc85>\d+)</td>\s*<td>\d+</td>\s*<td>\d+</td>\s*<td>[\d,.]+</td>\s*</tr>\s*" +
                            @"<tr[^>]*>\s*<td[^>]*>[^<]*</td>\s*<td>\d+</td>\s*<td>(?<pc75>\d+)</td>\s*<td>\d+</td>\s*<td>\d+</td>\s*<td>[\d,.]+</td>\s*</tr>\s*" +
                            @"<tr[^>]*>\s*<td[^>]*>[^<]*</td>\s*<td>\d+</td>\s*<td>(?<pc50>\d+)</td>\s*<td>\d+</td>\s*<td>\d+</td>\s*<td>[\d,.]+</td>\s*</tr>\s*" +
                            @"<tr[^>]*>\s*<td[^>]*>[^<]*</td>\s*<td>\d+</td>\s*<td>(?<NM>\d+)</td>\s*<td>\d+</td>\s*<td>\d+</td>\s*<td>[\d,.]+</td>\s*</tr>\s*</table><hr>";
                        break;
                    case FileTypes.StudioXML:
                        rgx = @"<file name=""(?<fullpath>[^""""]*)"" guid=""[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}"">\s*<analyse>\s*" +
                            @"<perfect segments=""\d+"" words=""(?<perfectsegments>\d+)"" characters=""\d+"" placeables=""\d+"" tags=""\d+"" />\s*" +
                            @"<inContextExact segments=""\d+"" words=""(?<inContextExact>\d+)"" characters=""\d+"" placeables=""\d+"" tags=""\d+"" />\s*" +
                            @"<exact segments=""\d+"" words=""(?<exactsegments>\d+)"" characters=""\d+"" placeables=""\d+"" tags=""\d+"" />\s*" +
                            @"<locked segments=""\d+"" words=""(?<lockedsegments>\d+)"" characters=""\d+"" placeables=""\d+"" tags=""\d+"" />\s*" +
                            @"<crossFileRepeated segments=""\d+"" words=""(?<crossFileRepeated>\d+)"" characters=""\d+"" placeables=""\d+"" tags=""\d+"" />\s*" +
                            @"<repeated segments=""\d+"" words=""(?<repeatedsegments>\d+)"" characters=""\d+"" placeables=""\d+"" tags=""\d+"" />\s*" +
                            @"<total segments=""\d+"" words=""(?<Total>\d+)"" characters=""\d+"" placeables=""\d+"" tags=""\d+"" />\s*" +
                            @"<new segments=""\d+"" words=""(?<newsegments>\d+)"" characters=""\d+"" placeables=""\d+"" tags=""\d+"" />\s*" +
                            @"<fuzzy min=""50"" max=""74"" segments=""\d+"" words=""(?<fuzzy1>\d+)"" characters=""\d+"" placeables=""\d+"" tags=""\d+"" />\s*" +
                            @"<fuzzy min=""75"" max=""84"" segments=""\d+"" words=""(?<fuzzy2>\d+)"" characters=""\d+"" placeables=""\d+"" tags=""\d+"" />\s*" +
                            @"<fuzzy min=""85"" max=""94"" segments=""\d+"" words=""(?<fuzzy3>\d+)"" characters=""\d+"" placeables=""\d+"" tags=""\d+"" />\s*" +
                            @"<fuzzy min=""95"" max=""99"" segments=""\d+"" words=""(?<fuzzy4>\d+)"" characters=""\d+"" placeables=""\d+"" tags=""\d+"" />\s*" +
                            @"</analyse>\s*</file>";
                        break;
                    case FileTypes.StudioHTML:
                        rgx = @"\s*<tr>\s*<th class=""File""[^>]*>(?!Files:\d+)(?<fullpath>[^<>]+)</th>\s*" +
                            @"<th[^>]*>PerfectMatch</th>\s*<td>\d+</td>\s*<td>(?<perfectsegments>\d+)</td>\s*<td>\d+</td>\s*<td>[\d+\.%]+</td>\s*<td>\d+</td>\s*<td>\d+</td>\s*</tr>\s*<tr>\s*<th[^>]*>Chars/Word:[\d\.]+</th>\s*" +
                            @"<th[^>]*>Context Match</th>\s*<td>\d+</td>\s*<td>(?<inContextExact>\d+)</td>\s*<td>\d+</td>\s*<td>[\d+\.%]+</td>\s*<td>\d+</td>\s*<td>\d+</td>\s*</tr>\s*<tr>\s*" +
                            @"<th[^>]*>Repetitions</th>\s*<td>\d+</td>\s*<td>(?<repeatedsegments>\d+)</td>\s*<td>\d+</td>\s*<td>[\d+\.%]+</td>\s*<td>\d+</td>\s*<td>\d+</td>\s*</tr>\s*<tr>\s*" +
                            @"<th[^>]*>Cross-file Repetitions</th>\s*<td>\d+</td>\s*<td>(?<crossFileRepeated>\d+)</td>\s*<td>\d+</td>\s*<td>[\d+\.%]+</td>\s*<td>\d+</td>\s*<td>\d+</td>\s*</tr>\s*<tr>\s*" +
                            @"<th[^>]*>100%</th>\s*<td>\d+</td>\s*<td>(?<exactsegments>\d+)</td>\s*<td>\d+</td>\s*<td>[\d+\.%]+</td>\s*<td>\d+</td>\s*<td>\d+</td>\s*</tr>\s*<tr>\s*" +
                            @"<th[^>]*>95% - 99%</th>\s*<td>\d+</td>\s*<td>(?<fuzzy4>\d+)</td>\s*<td>\d+</td>\s*<td>[\d+\.%]+</td>\s*<td>\d+</td>\s*<td>\d+</td>\s*</tr>\s*<tr>\s*" +
                            @"<th[^>]*>85% - 94%</th>\s*<td>\d+</td>\s*<td>(?<fuzzy3>\d+)</td>\s*<td>\d+</td>\s*<td>[\d+\.%]+</td>\s*<td>\d+</td>\s*<td>\d+</td>\s*</tr>\s*<tr>\s*" +
                            @"<th[^>]*>75% - 84%</th>\s*<td>\d+</td>\s*<td>(?<fuzzy2>\d+)</td>\s*<td>\d+</td>\s*<td>[\d+\.%]+</td>\s*<td>\d+</td>\s*<td>\d+</td>\s*</tr>\s*<tr>\s*" +
                            @"<th[^>]*>50% - 74%</th>\s*<td>\d+</td>\s*<td>(?<fuzzy1>\d+)</td>\s*<td>\d+</td>\s*<td>[\d+\.%]+</td>\s*<td>\d+</td>\s*<td>\d+</td>\s*</tr>\s*<tr>\s*" +
                            @"<th[^>]*>New</th>\s*<td>\d+</td>\s*<td>(?<newsegments>\d+)</td>\s*<td>\d+</td>\s*<td>[\d+\.%]+</td>\s*<td>\d+</td>\s*<td>\d+</td>\s*</tr>\s*<tr>\s*" +
                            @"<th[^>]*>Total</th>\s*<td[^>]*>\d+</td>\s*<td[^>]*>(?<Total>\d+)</td>\s*<td[^>]*>\d+</td>\s*<td[^>]*>[\d+\.%]+</td>\s*<td[^>]*>\d+</td>\s*<td[^>]*>\d+</td>\s*</tr>";
                        break;
                    default:
                        throw new Exception("Incorrect file type.");
                }


                MatchCollection matches = Regex.Matches(fullFile, rgx, RegexOptions.Singleline);
                foreach (Match match in matches)
                {
                    int[] wCounts;
                    int total = 0;
                    string fullPath = match.Groups["fullpath"].Value;
                    if (fullPath.Length == 0 && fileType == FileTypes.memoQHTML)
                    {
                        fullPath = match.Groups["filename"].Value;
                    }
                    if (translationFiles.ContainsKey(fullPath))
                    {
                        string tmpFullPath = fullPath;
                        int x = 1;
                        while (translationFiles.ContainsKey(tmpFullPath))
                        {
                            tmpFullPath = fullPath + x.ToString();
                            x++;
                        }
                        fullPath = tmpFullPath;
                    }

                    switch (fileType)
                    {
                        case FileTypes.memoQHTML:
                            wCounts = new int[9];
                            wCounts[0] = Convert.ToInt32(match.Groups["XTrans"].Value);
                            wCounts[1] = Convert.ToInt32(match.Groups["CM"].Value);
                            wCounts[2] = Convert.ToInt32(match.Groups["Reps"].Value);
                            wCounts[3] = Convert.ToInt32(match.Groups["pc100"].Value);
                            wCounts[4] = Convert.ToInt32(match.Groups["pc95"].Value);
                            wCounts[5] = Convert.ToInt32(match.Groups["pc85"].Value);
                            wCounts[6] = Convert.ToInt32(match.Groups["pc75"].Value);
                            wCounts[7] = Convert.ToInt32(match.Groups["pc50"].Value);
                            wCounts[8] = Convert.ToInt32(match.Groups["NM"].Value);
                            break;
                        case FileTypes.StudioHTML:
                        case FileTypes.StudioXML:
                            wCounts = new int[11];
                            if (match.Groups["lockedsegments"] != null && match.Groups["lockedsegments"].Value != "")
                            {
                                wCounts[0] = Convert.ToInt32(match.Groups["lockedsegments"].Value);
                            }
                            else
                            {
                                wCounts[0] = 0;
                            }
                            wCounts[1] = Convert.ToInt32(match.Groups["perfectsegments"].Value);
                            wCounts[2] = Convert.ToInt32(match.Groups["inContextExact"].Value);
                            wCounts[3] = Convert.ToInt32(match.Groups["repeatedsegments"].Value);
                            wCounts[4] = Convert.ToInt32(match.Groups["crossFileRepeated"].Value);
                            wCounts[5] = Convert.ToInt32(match.Groups["exactsegments"].Value);
                            wCounts[6] = Convert.ToInt32(match.Groups["fuzzy4"].Value);
                            wCounts[7] = Convert.ToInt32(match.Groups["fuzzy3"].Value);
                            wCounts[8] = Convert.ToInt32(match.Groups["fuzzy2"].Value);
                            wCounts[9] = Convert.ToInt32(match.Groups["fuzzy1"].Value);
                            wCounts[10] = Convert.ToInt32(match.Groups["newsegments"].Value);
                            break;
                        default:
                            throw new InvalidDataException("Invalid file format.");
                    }

                    total = Convert.ToInt32(match.Groups["Total"].Value);
                    TranslationFile tf = new TranslationFile(fullPath, wCounts, total, grid);
                    translationFiles.Add(fullPath, tf);

                }
            }


        }

        private void ParsememoQText()
        {
            throw new NotImplementedException("mQText not implemented");

        }

        private void CreateXLS()
        {
            string fileNamewoExt = Path.GetFileNameWithoutExtension(logfile);
            string fileNamewExt = fileNamewoExt + ".xlsx";
            string newXLSPath = Path.Combine(Path.GetDirectoryName(logfile), fileNamewExt);
            int i = 1;
            while (File.Exists(newXLSPath))
            {
                newXLSPath = Path.Combine(Path.GetDirectoryName(logfile), fileNamewoExt + i.ToString() + ".xlsx");
                i++;
            }

            Workbook wb = new Workbook();
            string wsName = fileNamewoExt.Length > 31 ? fileNamewoExt.Substring(0, 31) : fileNamewoExt;
            if (wb.Worksheets.Count == 0)
            {
                ws = wb.Worksheets.Add(wsName);
            }
            else
            {
                ws = wb.Worksheets[0];
                ws.Name = wsName;
            }

            writeWSHeader();
            Program.mainWindow.updateProgress(40);

            writeWSData();
            Program.mainWindow.updateProgress(70);

            Helper.formatColumns(ref ws);
            Program.mainWindow.updateProgress(90);

            ws.FreezePanes(3, 2, 3, 2);
            Program.mainWindow.updateProgress(100);

            wb.Save(newXLSPath);
            System.Diagnostics.Process.Start(newXLSPath);

        }

        private void writeWSHeader()
        {
            string[] headerTitles;
            int[] formulaChooser;
            int[] isMerge;

            switch (fileType)
            {
                case FileTypes.memoQ:
                case FileTypes.memoQAllInfo:
                case FileTypes.memoQHTML:
                case FileTypes.memoQTrados:
                    headerTitles = new string[]{ "Path", "File name", "X-translated", "Context TM", "Repetitions", "100% Matches", "95% - 99%", "85% - 94%",
                        "75% - 84%", "50% - 74%", "No Match", "Total", "Weighted", "DTP sent", "DTP provider", "Deadline", "Ready",
                        "Imported to memoQ", "Ready to translate", "Translator", "Deadline", "Ready", "Reviewer", "Deadline", "Ready", "Ready to deliver" };

                    formulaChooser = new int[] { 1, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 3, 1, 1, 3, 3, 3, 1, 1, 3, 1, 1, 3, 3 };
                    isMerge = new int[] { 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 };
                    TotalColumn = 12;
                    columnColors = new Color[] { c5, c5, c6, c6, c6, c6, c6, c6, c6, c6, c6, c7, c7, c1, c1, c1, c1, c2, c2, c3, c3, c2, c3, c3, c2, c4 };
                    Program.ColumnWidth = new int[] { -1, -1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 100, 100, 100, 120, 120, 100, 100, 100, 120, 120, 100, 120, 120, 100, 100 };
                    Program.isCondForm = new int[] { 1, 0, 0, 1, 1, 1, 0, 0, 1, 0, 0, 1, 1 };
                    break;
                case FileTypes.StudioHTML:
                case FileTypes.StudioXML:
                    headerTitles = new string[]{ "Path", "File name", "Locked", "PerfectMatch", "Context Match", "Repetitions", "Cross-file Repetitions", "100% Matches", "95% - 99%", "85% - 94%",
                        "75% - 84%", "50% - 74%", "No Match", "Total", "Weighted", "DTP sent", "DTP provider", "Deadline", "Ready",
                        "Imported to memoQ", "Ready to translate", "Translator", "Deadline", "Ready", "Reviewer", "Deadline", "Ready", "Ready to deliver" };

                    formulaChooser = new int[] { 1, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 3, 1, 1, 3, 3, 3, 1, 1, 3, 1, 1, 3, 3 };
                    isMerge = new int[] { 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 };
                    TotalColumn = 14;
                    columnColors = new Color[] { c5, c5, c6, c6, c6, c6, c6, c6, c6, c6, c6, c6, c6, c7, c7, c1, c1, c1, c1, c2, c2, c3, c3, c2, c3, c3, c2, c4 };
                    Program.ColumnWidth = new int[] { -1, -1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 100, 100, 100, 120, 120, 100, 100, 100, 120, 120, 100, 120, 120, 100, 100 };
                    Program.isCondForm = new int[] { 1, 0, 0, 1, 1, 1, 0, 0, 1, 0, 0, 1, 1 };
                    grid = Helper.GridConvertToStudio(grid);
                    break;
                default:
                    throw new InvalidDataException("Invalid file type.");
            }
            for (int i = 0; i < TotalColumn - 3; i++)
            {
                Cell c = ws.Cells[0, 2 + i];
                c.Value = Convert.ToDouble(grid[i]) / 100;
                Helper.AddBold(c);
                Helper.AddBackground(c, columnColors[2 + i]);
                Helper.AddNumberFormat(c, "0%");

            }

            int columncount = headerTitles.Length;
            string lastrow = (translationFiles.Count + 3).ToString();

            for (int i = 0; i < formulaChooser.Length; i++)
            {
                if (isMerge[i] == 1)
                {
                    Aspose.Cells.Range r = ws.Cells.CreateRange(0, i, 2, 1);
                    r.Merge();
                }

                Cell c = ws.Cells[1 - isMerge[i], i];
                c.Value = headerTitles[i];
                Helper.AddBold(c);
                Helper.AddBackground(c, columnColors[i]);
                Helper.AddAlignment(c);

                string formula = "";
                string column = CellsHelper.ColumnIndexToName(i);

                switch (formulaChooser[i])
                {
                    case 1:
                        formula = "=COUNTA($" + column + "$4:$" + column + "$" + lastrow + ")";
                        break;
                    case 2:
                        formula = "=SUBTOTAL(9, $" + column + "$4:$" + column + "$" + lastrow + ")";
                        break;
                    case 3:
                        formula = "=COUNTIF($" + column + "$4:$" + column + "$" + lastrow + ",\"x\")";
                        break;
                }
                c = ws.Cells[2, i];
                c.Formula = formula;
                Helper.AddBold(c);
                Helper.AddBackground(c, columnColors[i]);

            }
            ws.Cells[2, 0].Value = "Total:";
            Helper.AddNumberFormat(ws.Cells[2, TotalColumn], "0");
        }

        private void writeWSData()
        {
            for (int sor = 0; sor < translationFiles.Count; sor++)
            {
                TranslationFile tf = translationFiles.ElementAt(sor).Value;
                Cell c = ws.Cells[3 + sor, 0];
                c.Value = tf.RelativePath;

                c = ws.Cells[3 + sor, 1];
                c.Value = tf.FileName;

                for (int oszlop = 0; oszlop < TotalColumn - 3; oszlop++)
                {
                    ws.Cells[3 + sor, 2 + oszlop].Value = tf.wCounts[oszlop];
                }

                string lastWCcolumn = CellsHelper.ColumnIndexToName(TotalColumn - 2);
                c = ws.Cells[3 + sor, TotalColumn - 1];
                c.Formula = "=SUM(C" + (4 + sor).ToString() + ":" + lastWCcolumn + (4 + sor).ToString() + ")";
                Helper.AddBold(c);

                c = ws.Cells[3 + sor, TotalColumn];
                c.Formula = "=SUMPRODUCT(C" + (4 + sor).ToString() + ":" + lastWCcolumn + (4 + sor).ToString() + ",C$1:" + lastWCcolumn + "$1)";
                Helper.AddBold(c);
                Helper.AddNumberFormat(c, "0");

                for (int oszlop = TotalColumn + 1; oszlop < TotalColumn + 14; oszlop++)
                {
                    c = ws.Cells[3 + sor, oszlop];
                    Helper.AddBackground(c, columnColors[oszlop]);
                    if (oszlop == TotalColumn + 3 || oszlop == TotalColumn + 8 || oszlop == TotalColumn + 11)
                    {
                        Helper.AddNumberFormat(c, "mmmm d.");
                    }
                }

            }

            Helper.AddConditionalFormatting(ref ws, translationFiles.Count + 2, TotalColumn);
        }


    }
}
