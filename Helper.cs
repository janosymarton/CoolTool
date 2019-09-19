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
using ResourcesAPI_TBList;

namespace CoolTool
{
    static class Helper
    {
        public static bool FindSeparator(string headerline, out char Separator, out bool Success)
        {
            Success = true;
            if (headerline.Contains(ControlChar.TabChar))
            {
                Separator = ControlChar.TabChar;
                return true;
            }
            else if (headerline.Contains(';'))
            {
                Separator = ';';
                return true;
            }
            else
            {
                Success = false;
                throw new Exception("csv separator cannot be found. Invalid logfile format.");
            }

        }


        public static void FindCommonPath(ref Dictionary<string, TranslationFile> translationFiles)
        {
            string commonPath = "";
            bool commonPathInit = true;
            foreach (string tfFullPath in translationFiles.Keys)
            {
                string fullpath = "";
                try
                {
                    fullpath = Path.GetDirectoryName(translationFiles[tfFullPath].FullPath);
                }
                catch
                {
                    commonPath = "";
                    Log.AddLog("File without path found.");
                    break;
                }

                string onlyPath = fullpath.Length > 0 ? fullpath + "\\" : "";
                if (commonPathInit && commonPath == "")
                {
                    commonPath = onlyPath;
                    commonPathInit = false;
                }
                else if (!commonPathInit && commonPath != "")
                {
                    int i = commonPath.Length - 1;
                    while (i > 0 && !onlyPath.Contains(commonPath.Substring(0, i)))
                    {
                        i--;
                    }

                    commonPath = commonPath.Substring(0, i + 1);
                }
            }

            if (commonPath.Contains('\\'))
            {
                commonPath = commonPath.Substring(0, commonPath.LastIndexOf('\\'));
                foreach (string tfFullPath in translationFiles.Keys)
                {
                    string fp = Path.GetDirectoryName(translationFiles[tfFullPath].FullPath);
                    translationFiles[tfFullPath].RelativePath =
                        fp.Substring(commonPath.Length, fp.Length - commonPath.Length);
                }
            }
            else
            {
                commonPath = "";
            }
        }

        public static void AddBackground(Cell c, Color color)
        {
            Aspose.Cells.Style style = c.GetStyle();
            style.Pattern = BackgroundType.Solid;
            style.ForegroundColor = color;
            c.SetStyle(style);
        }
        public static void AddBackground(Cell c, Color color, BackgroundType type)
        {
            Aspose.Cells.Style style = c.GetStyle();
            style.Pattern = type;
            style.ForegroundColor = color;
            c.SetStyle(style);
        }


        public static void AddBold(Cell c)
        {
            Aspose.Cells.Style style = c.GetStyle();
            style.Font.IsBold = true;
            c.SetStyle(style);
        }

        public static void AddFontColor(Cell c, Color color)
        {
            Aspose.Cells.Style style = c.GetStyle();
            style.Font.Color = color;
            c.SetStyle(style);
        }


        public static void AddNumberFormat(Cell c, string numberFormat)
        {
            Aspose.Cells.Style style = c.GetStyle();
            style.Custom = numberFormat;
            c.SetStyle(style);
        }

        public static void AddAlignment(Cell c)
        {
            Aspose.Cells.Style style = c.GetStyle();
            style.HorizontalAlignment = TextAlignmentType.Center;
            style.VerticalAlignment = TextAlignmentType.Center;
            style.IsTextWrapped = true;
            c.SetStyle(style);
        }

        public static void AddConditionalFormatting(ref Worksheet ws, int endRow, int TotalColumn)
        {
            int index = ws.ConditionalFormattings.Add();
            FormatConditionCollection conds = ws.ConditionalFormattings[index];
            int[] isCondForm = Program.isCondForm;

            for (int i = 0; i < isCondForm.Length; i++)
            {
                if (isCondForm[i] == 1)
                {
                    CellArea area = new CellArea();
                    area.StartRow = 3;
                    area.EndRow = endRow;
                    area.StartColumn = TotalColumn + 1 + i;
                    area.EndColumn = TotalColumn + 1 + i;
                    conds.AddArea(area);
                }
            }

            int idx = conds.AddCondition(FormatConditionType.ContainsText);
            FormatCondition cond = conds[idx];
            cond.Style.BackgroundColor = Color.LightGreen;
            cond.Style.Pattern = BackgroundType.Solid;
            cond.Text = "x";
        }

        public static void formatColumns(ref Worksheet ws)
        {
            int[] columnWidth = Program.ColumnWidth;
            for (int i = 0; i < columnWidth.Length; i++)
            {
                switch (columnWidth[i])
                {
                    case -1:
                        ws.AutoFitColumn(i);
                        break;
                    case 0:
                        ws.Cells.SetColumnWidthPixel(i, 100);
                        ws.Cells.HideColumn(i);
                        break;
                    default:
                        ws.Cells.SetColumnWidthPixel(i, columnWidth[i]);
                        break;

                }
            }
        }

        public static void autoFitColumns(ref Worksheet ws)
        {
            int column = 0;
            while (ws.Cells[0, column].Value != null)
            {
                ws.AutoFitColumn(column);
                column++;
            }
        }

        internal static int[] GridConvertToStudio(int[] grid)
        {
            int[] tmpGrid =
                new int[] {
                            0,
                            grid[0],
                            grid[1],
                            grid[2],
                            grid[2],
                            grid[3],
                            grid[4],
                            grid[5],
                            grid[6],
                            grid[7],
                            grid[8] };
            return tmpGrid;
        }

        internal static Dictionary<Guid, string> TBList2stringArray(TBListResponse[] origTBlist)
        {
            Dictionary<Guid, string> tbList = new Dictionary<Guid, string>();

            foreach (TBListResponse tbItem in origTBlist)
            {
                tbList.Add(new Guid(tbItem.TBGuid), tbItem.FriendlyName);
            }

            var ordered = tbList.OrderBy(x => x.Value);
            Dictionary<Guid, string> orderedDict = new Dictionary<Guid, string>();

            foreach (var item in ordered)
            {
                orderedDict.Add(item.Key, item.Value);
            }

            return orderedDict;
        }
    }
}
