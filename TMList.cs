using System;
using ResourcesAPI_TMList;
using Aspose.Cells;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace CoolTool
{
    internal class TMList
    {
        private string exportFile;
        private string client;

        public TMList(string file, string text, int mode)
        {
            this.exportFile = file;
            this.client = text;
            
            if (mode == 0)
            {
                getSimpleTMlist();
            }
            else if (mode == 1)
            {
                getMergeTMlist();
            }
            
        }

        private void getSimpleTMlist()
        {
            TMListResponse[] TMlist = RAPI_Session.generateTMlist();

            Program.mainWindow.updateProgress(50);

            Worksheet ws;
            Workbook wb = new Workbook();
            string wsName = "TM list";
            if (wb.Worksheets.Count == 0)
            {
                ws = wb.Worksheets.Add(wsName);
            }
            else
            {
                ws = wb.Worksheets[0];
                ws.Name = wsName;
            }

            WriteHeader(ref ws);
            int rowCount = 1;

            foreach (TMListResponse tmListObj in TMlist)
            {
                if (client == "***All***" || tmListObj.Client == client)
                {
                    WriteTM(tmListObj, ref ws, rowCount);
                    rowCount++;
                }
            }

            Helper.autoFitColumns(ref ws);
            wb.Save(exportFile);
            Program.mainWindow.updateProgress(100);
        }

        private void getMergeTMlist()
        {
            LanguageHelper.Initialize();

            TMListResponse[] TMlist = RAPI_Session.generateTMlist();
            Dictionary<string[], ClientTMs> TMGroups = new Dictionary<string[], ClientTMs>();
            Dictionary<string[], ClientTMs> sortedTMGroups = new Dictionary<string[], ClientTMs>();
            Program.mainWindow.updateProgress(25);

            foreach (TMListResponse tm in TMlist)
            {
                if (tm.Client.Length > 0 && (client == "***All***" || tm.Client == client))
                {
                    string[] newid = new string[] { tm.Client, tm.SourceLangCode, tm.TargetLangCode };
                    string[] existID = null;

                    foreach (string[] id in TMGroups.Keys)
                    {
                        if (tm.Client == TMGroups[id].client && tm.SourceLangCode == TMGroups[id].sLang && tm.TargetLangCode == TMGroups[id].tLang)
                        {
                            existID = id;
                        }
                    }

                    if (existID == null)
                    {
                        ClientTMs newGroup = new ClientTMs(newid);
                        TMGroups.Add(newid, newGroup);
                        TMGroups[newid].TMlist.Add(tm);
                    }
                    else
                    {
                        TMGroups[existID].TMlist.Add(tm);
                    }
                }
            }

            Program.mainWindow.updateProgress(50);
            
            IOrderedEnumerable<KeyValuePair<string[], ClientTMs>> linqGroups = from clientTMgroup in TMGroups
                        orderby clientTMgroup.Key[0], clientTMgroup.Key[1], clientTMgroup.Key[2] ascending
                        select clientTMgroup;

            
            Worksheet ws;
            Workbook wb = new Workbook();
            string wsName = "TM list";
            if (wb.Worksheets.Count == 0)
            {
                ws = wb.Worksheets.Add(wsName);
            }
            else
            {
                ws = wb.Worksheets[0];
                ws.Name = wsName;
            }

            WriteHeader(ref ws);

            int rowCount = 0;

            Program.mainWindow.updateProgress(75);

            foreach (KeyValuePair<string[], ClientTMs> kvp in linqGroups)
            {
                sortedTMGroups.Add(kvp.Key, kvp.Value);
            }

            foreach(string[] id in sortedTMGroups.Keys)
            { 
                if (sortedTMGroups[id].TMlist.Count > 0)
                {
                    rowCount++;
                    Cell c = ws.Cells[rowCount, 0];
                    c.Value = sortedTMGroups[id].client + " (" + sortedTMGroups[id].sLang + " > " + sortedTMGroups[id].tLang + ")";
                    Helper.AddFontColor(c, Color.Purple);

                    rowCount++;

                    sortedTMGroups[id].findMasterTM();

                    TMListResponse tm = sortedTMGroups[id].getMasterTM();
                    WriteTM(tm, ref ws, rowCount);
                    FormatLine(tm, rowCount, ref ws);

                    rowCount++;

                    TMListResponse[] resultList = sortedTMGroups[id].getOtherTM();
                    foreach (TMListResponse tmn in resultList)
                    {
                        WriteTM(tmn, ref ws, rowCount);
                        FormatLine(tmn, rowCount, ref ws);
                        rowCount++;
                    }

                }
            }

            Helper.autoFitColumns(ref ws);
            wb.Save(exportFile);
            Program.mainWindow.updateProgress(100);

        }

        private static void FormatLine(TMListResponse tm, int rowCount, ref Worksheet ws)
        {
            Color color = Color.White;
            BackgroundType type = BackgroundType.Solid;

            if (tm.NumEntries == 0)
            {
                color = Color.LightGray;
                type = BackgroundType.DiagonalStripe;
            }

            if (tm.FriendlyName.ToLower().Contains("master"))
            {
                color = Color.LightPink;
            }

            int colCount = 0;

            while (ws.Cells[0, colCount].Value != null)
            {
                Cell c = ws.Cells[rowCount, colCount];
                Helper.AddBackground(c, color, type);
                colCount++;
            }

        }

        private static void WriteHeader(ref Worksheet ws)
        {
            Cell c;
            c = ws.Cells[0, 0];
            c.Value = "Guid";
            Helper.AddBold(c);

            c = ws.Cells[0, 1];
            c.Value = "Name";
            Helper.AddBold(c);

            c = ws.Cells[0, 2];
            c.Value = "Source";
            Helper.AddBold(c);

            c = ws.Cells[0, 3];
            c.Value = "Target";
            Helper.AddBold(c);

            c = ws.Cells[0, 4];
            c.Value = "Client";
            Helper.AddBold(c);

            c = ws.Cells[0, 5];
            c.Value = "Domain";
            Helper.AddBold(c);

            c = ws.Cells[0, 6];
            c.Value = "Subject";
            Helper.AddBold(c);

            c = ws.Cells[0, 7];
            c.Value = "Project";
            Helper.AddBold(c);

            c = ws.Cells[0, 8];
            c.Value = "Size";
            Helper.AddBold(c);

            c = ws.Cells[0, 9];
            c.Value = "TM Owner";
            Helper.AddBold(c);

        }

        private static void WriteTM(TMListResponse tmListObj, ref Worksheet ws, int rowCount)
        {
            Cell c;
            c = ws.Cells[rowCount, 0];
            c.Value = tmListObj.TMGuid;

            c = ws.Cells[rowCount, 1];
            c.Value = tmListObj.FriendlyName;

            c = ws.Cells[rowCount, 2];
            c.Value = tmListObj.SourceLangCode;

            c = ws.Cells[rowCount, 3];
            c.Value = tmListObj.TargetLangCode;

            c = ws.Cells[rowCount, 4];
            c.Value = tmListObj.Client;

            c = ws.Cells[rowCount, 5];
            c.Value = tmListObj.Domain;

            c = ws.Cells[rowCount, 6];
            c.Value = tmListObj.Subject;

            c = ws.Cells[rowCount, 7];
            c.Value = tmListObj.Project;

            c = ws.Cells[rowCount, 8];
            c.Value = tmListObj.NumEntries;

            c = ws.Cells[rowCount, 9];
            c.Value = tmListObj.TMOwner;
        }


    }
}