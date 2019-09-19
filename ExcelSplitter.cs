using Aspose.Cells;
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace CoolTool
{
    internal class ExcelSplitter
    {
        private string XLSpath;


        public ExcelSplitter(string file, int excelMode)
        {
            this.XLSpath = file;

            try
            {
                Program.mainWindow.updateProgress(0);

                switch (excelMode)
                {
                    case 1:
                        RunJoinerHide(); // Join
                        break;
                    case 2:
                        RunSplitterHide(); // Split
                        break;
                    case 3:
                        RunPDFSaver(); // PDF
                        break;
                    case 4:
                        RunStyleChanger();
                        break;
                }

                Program.mainWindow.updateProgress(100);
                Log.AddLog("Processing finished");
            }
            catch (Exception ex)
            {
                Log.AddLog("Problem while processing: " + ex.Message, true);
            }
        }

        private void RunStyleChanger()
        {
            string filePath = Path.GetDirectoryName(XLSpath);
            string fileName = Path.GetFileNameWithoutExtension(XLSpath);
            string tmpFile = Path.Combine(filePath, "backup_" + fileName + Path.GetExtension(XLSpath));

            File.Copy(XLSpath, tmpFile);

            Workbook wb = new Workbook(XLSpath);
            WorksheetCollection wsc = wb.Worksheets;
            int i = 1;
            int noOfSheets = wsc.Count;

            for (int wsIndex = 0; wsIndex < noOfSheets; wsIndex++)
            {
                try
                {
                    Worksheet ws = wb.Worksheets[wsIndex];
                    StyleChanger(ref ws);
                    i++;
                    Program.mainWindow.updateProgress(Convert.ToInt32(100.0 * Convert.ToDouble(i) / Convert.ToDouble(noOfSheets)));
                }
                catch (Exception ex)
                {
                    Log.AddLog("Error while converting sheet " + wsIndex.ToString() + " - " + ex.Message, true);
                }
            }

            wb.Save(XLSpath);
            Log.AddLog("Styles changed in " + XLSpath);
        }

        private void RunPDFSaver()
        {
            string filePath = Path.GetDirectoryName(XLSpath);
            string fileName = Path.GetFileNameWithoutExtension(XLSpath);
            string tmpFile = Path.Combine(filePath, "tmp_" + fileName + Path.GetExtension(XLSpath));

            File.Copy(XLSpath, tmpFile);

            Workbook wb = new Workbook(tmpFile);
            WorksheetCollection wsc = wb.Worksheets;
            int i = 1;
            int noOfSheets = wsc.Count;
            int noOfSheetsDigits = noOfSheets.ToString().Length;

            for (int wsIndex = 1; wsIndex < noOfSheets; wsIndex++)
            {
                wb.Worksheets[wsIndex].IsVisible = false;
            }

            for (int wsIndex = 0; wsIndex < noOfSheets; wsIndex++)
            {
                try
                {
                    Worksheet ws = wb.Worksheets[wsIndex];
                    StyleChanger(ref ws);

                    string newFileName = ws.Name + "_" + fileName + ".pdf";
                    string newFilePath = Path.Combine(filePath, newFileName);
                    PdfSaveOptions pso = new PdfSaveOptions();
                    pso.OnePagePerSheet = true;
                    pso.CalculateFormula = true;
                    wb.Save(newFilePath, pso);

                    if (wsIndex < noOfSheets - 1)
                    {
                        wb.Worksheets[wsIndex + 1].IsVisible = true;
                        wb.Worksheets[wsIndex].IsVisible = false;
                    }

                    i++;
                    Program.mainWindow.updateProgress(Convert.ToInt32(100.0 * Convert.ToDouble(i) / Convert.ToDouble(noOfSheets)));
                }
                catch (Exception ex)
                {
                    Log.AddLog("Error while converting sheet " + wsIndex.ToString() + " - " + ex.Message, true);
                }
            }

            //wb.Save(tmpFile);

            File.Delete(tmpFile);
        }

        private void StyleChanger(ref Worksheet ws)
        {
            int rowmin = 7;
            int rowmax = 44;
            int columnmin = 1;
            int columnmax = ws.Cells.MaxColumn;
            if (columnmax < 4) columnmax = 4;

            for (int i = 0; i < columnmax; i++)
            {
                double columnWidth = ws.Cells.Columns[i].Width;
                columnWidth = columnWidth * 1.15;
                ws.Cells.Columns[i].Width = columnWidth;
            }

            for (int row = rowmin - 1; row < rowmax; row++)
            {
                for (int column = columnmin - 1; column < columnmax; column++)
                {
                    Cell c = ws.Cells[row, column];
                    Style style = c.GetStyle();
                    style.IsTextWrapped = true;
                    c.SetStyle(style);
                }
                ws.AutoFitRow(row);
            }

            ws.PageSetup.PrintArea = "A1:" + (char)(ws.Cells.MaxDataColumn + 'A') + (ws.Cells.MaxDataRow + 1);
        }

        private void RunSplitterHide()
        {

            Workbook wb = new Workbook(XLSpath);
            WorksheetCollection wsc = wb.Worksheets;
            string filePath = Path.GetDirectoryName(XLSpath);
            string fileName = Path.GetFileNameWithoutExtension(XLSpath);
            int i = 1;
            int noOfSheets = wsc.Count;
            int noOfSheetsDigits = noOfSheets.ToString().Length;

            foreach (Worksheet ws in wsc)
            {
                int wsIndex = ws.Index;
                string newFileName = "Sheet_" + SeqNumber(i, noOfSheetsDigits) + "_" + fileName + "_" + ws.Name + ".xlsx";
                string newFilePath = Path.Combine(filePath, newFileName);
                File.Copy(XLSpath, newFilePath);
                Workbook newWB = new Workbook(newFilePath);
                foreach (Worksheet wsInNewFile in newWB.Worksheets)
                {
                    if (wsInNewFile.Index != wsIndex)
                    {
                        wsInNewFile.IsVisible = false;
                    }
                }

                newWB.Save(newFilePath, SaveFormat.Xlsx);
                i++;
                Program.mainWindow.updateProgress(Convert.ToInt32(100.0 * Convert.ToDouble(i) / Convert.ToDouble(noOfSheets)));
            }

        }

        private string SeqNumber(int i, int noOfSheetsDigits)
        {
            int CurrentLength = i.ToString().Length;
            string result = new String('0', noOfSheetsDigits - CurrentLength);
            result += i.ToString();
            return result;
        }

        private void RunJoinerHide()
        {
            string filePath = Path.GetDirectoryName(XLSpath);
            string fileName = Path.GetFileNameWithoutExtension(XLSpath);
            string tmpWBName = "MERGED_" + fileName + ".xlsx";
            string tmpWBPath = Path.Combine(filePath, tmpWBName);

            File.Copy(XLSpath, tmpWBPath);

            Workbook wb = new Workbook(tmpWBPath);
            WorksheetCollection wsc = wb.Worksheets;

            int i = 1;
            int noOfSheets = wsc.Count;
            int noOfInvalidSheets = 0;
            int noOfSheetsDigits = noOfSheets.ToString().Length;
            foreach (Worksheet ws in wsc)
            {
                string newFileName = "Sheet_" + SeqNumber(i, noOfSheetsDigits) + "_" + fileName + "_" + ws.Name + ".xlsx";
                string newFilePath = Path.Combine(filePath, newFileName);

                if (File.Exists(newFilePath))
                {
                    try
                    {
                        Workbook wbToMerge = new Workbook(newFilePath);
                        WorksheetCollection wscToMerge = wbToMerge.Worksheets;
                        Worksheet wsToMerge = wscToMerge[ws.Index];

                        CopyOptions co = new CopyOptions();
                        co.CopyInvalidFormulasAsValues = false;
                        co.ReferToSheetWithSameName = true;
                        co.ReferToDestinationSheet = true;
                        ws.Copy(wsToMerge, co);
                    }
                    catch
                    {
                        noOfInvalidSheets++;
                    }
                }
                else
                {
                    Log.AddLog("File not found: " + newFilePath, true);
                    noOfInvalidSheets++;
                }
                i++;
                Program.mainWindow.updateProgress(Convert.ToInt32(100 * Convert.ToDouble(i) / Convert.ToDouble(noOfSheets)));
            }

            wb.Save(tmpWBPath, SaveFormat.Xlsx);
            if (noOfInvalidSheets > 0)
            {
                throw new Exception("There are " + noOfInvalidSheets.ToString() + " invalid sheets.");
            }
        }

        //private void RunSplitter()
        //{
        //    Workbook wb = new Workbook(XLSpath);
        //    WorksheetCollection wsc = wb.Worksheets;
        //    string filePath = Path.GetDirectoryName(XLSpath);
        //    string fileName = Path.GetFileNameWithoutExtension(XLSpath);
        //    int i = 1;
        //    int noOfSheets = wsc.Count;

        //    foreach (Worksheet ws in wsc)
        //    {
        //        string newFileName = "Sheet_" + i.ToString() + "_" + fileName + "_" + ws.Name + ".xlsx";
        //        string newFilePath = Path.Combine(filePath, newFileName);
        //        Workbook newWB = new Workbook();
        //        newWB.Worksheets[0].Name = ws.Name;
        //        newWB.Worksheets[0].Copy(ws);

        //        string FindText = filePath + @"\[" + fileName + ".xlsx]";
        //        Cell cell = null;
        //        FindOptions fo = new FindOptions();
        //        fo.CaseSensitive = false;
        //        fo.LookInType = LookInType.Formulas;
        //        fo.LookAtType = LookAtType.Contains;

        //        do
        //        {
        //            cell = newWB.Worksheets[0].Cells.Find(FindText, cell, fo);
        //            if (cell != null)
        //            {
        //                string celltext = cell.Formula.ToString();
        //                celltext = celltext.Replace(FindText, "");
        //                cell.Formula = celltext;
        //            }
        //        }
        //        while (cell != null);


        //        newWB.Save(newFilePath, SaveFormat.Xlsx);
        //        i++;
        //        mainWindow.updateProgress(Convert.ToInt32(100.0 * Convert.ToDouble(i) / Convert.ToDouble(noOfSheets)));
        //    }

        //}


        //private void RunJoiner()
        //{

        //    Workbook newWB = new Workbook();
        //    Workbook wb = new Workbook(XLSpath);
        //    WorksheetCollection wsc = wb.Worksheets;
        //    string filePath = Path.GetDirectoryName(XLSpath);
        //    string fileName = Path.GetFileNameWithoutExtension(XLSpath);
        //    int i = 1;
        //    int noOfSheets = wsc.Count;
        //    int noOfInvalidSheets = 0;
        //    foreach(Worksheet ws in wsc)
        //    {
        //        string newFileName = "Sheet_" + i.ToString() + "_" + fileName + "_" + ws.Name + ".xlsx";
        //        string newFilePath = Path.Combine(filePath, newFileName);
        //        if (File.Exists(newFilePath))
        //        {
        //            try
        //            {
        //                Workbook wbToMerge = new Workbook(newFilePath);
        //                WorksheetCollection wscToMerge = wbToMerge.Worksheets;
        //                Worksheet wsToMerge = wscToMerge[0];
        //                if (newWB.Worksheets.Count < i)
        //                {
        //                    newWB.Worksheets.Add(ws.Name);
        //                }
        //                else
        //                {
        //                    newWB.Worksheets[i - 1].Name = ws.Name;
        //                }

        //                Worksheet newWS = newWB.Worksheets[i - 1];
        //                CopyOptions co = new CopyOptions();
        //                co.CopyInvalidFormulasAsValues = false;
        //                newWS.Copy(wsToMerge, co);


        //            }
        //            catch
        //            {
        //                noOfInvalidSheets++;
        //            }
        //        }
        //        else
        //        {
        //            Log.AddLog("File not found: " + newFilePath, true);
        //            noOfInvalidSheets++;
        //        }
        //        i++;
        //        mainWindow.updateProgress(Convert.ToInt32(80.0 * Convert.ToDouble(i) / Convert.ToDouble(noOfSheets)));
        //    }

        //    i = 0;
        //    string tmpWBName = "MERGED_" + fileName + ".xlsx";
        //    string tmpWBPath = Path.Combine(filePath, tmpWBName);
        //    newWB.Save(tmpWBPath, SaveFormat.Xlsx);

        //    foreach (Worksheet wsn in newWB.Worksheets)
        //    {
        //        try
        //        {
        //            Cell cell = null;
        //            FindOptions fo = new FindOptions();
        //            fo.CaseSensitive = false;
        //            fo.LookInType = LookInType.Formulas;
        //            fo.LookAtType = LookAtType.Contains;

        //            do
        //            {
        //                cell = wsn.Cells.Find("[", cell, fo);
        //                if (cell != null)
        //                {
        //                    string celltext = cell.Formula.ToString();
        //                    Regex rgx = new Regex(@"(?<=')(([^\[']+)?\[(?<sheet>((?<truncSheet>[^\]]+)\.)?[^\]]+)\](\k<sheet>|\k<truncSheet>))(?=')");
        //                    celltext = rgx.Replace(celltext, "${sheet}");
        //                    cell.Formula = celltext;
        //                }
        //            }
        //            while (cell != null);
        //            i++;
        //            mainWindow.updateProgress(80 + Convert.ToInt32(20.0 * Convert.ToDouble(i) / Convert.ToDouble(noOfSheets)));
        //        }
        //        catch (Exception ex)
        //        {
        //            throw new Exception("Error: " + i.ToString() + " - " + ex.Message);
        //        }
        //    }

        //    newWB.Save(tmpWBPath, SaveFormat.Xlsx);
        //    if(noOfInvalidSheets>0)
        //    {
        //        throw new Exception("There are " + noOfInvalidSheets.ToString() + " invalid sheets.");
        //    }
        //}
    }
}