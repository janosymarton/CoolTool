using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace CoolTool
{
    class TranslationFile
    {
        public string FullPath;
        public string RelativePath;
        public string FileName;
        public int[] wCounts;
        public int Total;
        public double WWC;

        public TranslationFile(string FullPath, int[] wCounts, int Total, int[] grid)
        {
            this.wCounts = wCounts;
            this.Total = Total;
            string regex = @"(?<idezo>""?)(\[[\w\-_]+\]\s)?(?<filename>[A-Za-z]:\\.*[^""])\k<idezo>";
            if(Regex.IsMatch(FullPath, regex))
            {
                FullPath = Regex.Replace(FullPath, regex, "${filename}");
            }
            this.FullPath = FullPath;

            try
            {
                this.FileName = Path.GetFileName(FullPath);
            }
            catch
            {
                this.FileName = FullPath;
            }

            int calcTotal = 0;
            double calcWWC = 0;
            for (int i = 0; i < 9; i++)
            {
                calcTotal += wCounts[i];
                calcWWC += Convert.ToDouble(wCounts[i]) * grid[i] / 100;
            }

            this.WWC = calcWWC;

            if (Total != calcTotal)
            {
                Log.AddLog("Total wordcount is not matching for file: " + FullPath, true);
            }
        }

        public void UpdateRelativePath(string RelativePath)
        {
            this.RelativePath = RelativePath;
        }

    }
}
