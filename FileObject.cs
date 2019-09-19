
using System.IO;
using System.Text.RegularExpressions;

namespace CoolTool
{
    internal class FileObject
    {
        public string fileType;
        public string fullPath;
        public string filename;
        public string path;
        public int pageCount = 0;
        public int wordCount = 0;
        public int imageCount = 0;
        public int embeddedDocsCount = 0;
        public bool hasPassword = false;
        public string comment;
        public int characterCount = 0;

        public FileObject(string path)
        {
            this.fullPath = path;
            string ext = Path.GetExtension(path).ToLower();
            switch (ext)
            {
                case ".xlsx":
                case ".xls":
                case ".csv":
                    fileType = "Excel";
                    break;
                case ".docx":
                case ".doc":
                    fileType = "Word";
                    break;
                case ".pdf":
                    fileType = "PDF";
                    break;
                case ".jpg":
                case ".png":
                case ".bmp":
                case ".ico":
                case ".gif":
                    fileType = "Image";
                    break;
                case ".ppt":
                case ".pptx":
                    fileType = "PowerPoint";
                    break;
                case ".zip":
                    fileType = "ZIP";
                    break;
                case ".html":
                case ".htm":
                    fileType = "Webpage";
                    break;
                case ".indd":
                case ".idml":
                    fileType = "InDesign";
                    break;

                default:
                    fileType = "???";
                    break;
            }

            string rgx = @"(?m)(?<folder>^.+\\)(?<fajl>[^\\]+)$";
            MatchCollection matches = Regex.Matches(path, rgx);
            if (matches.Count == 1)
            {
                this.filename = matches[0].Groups["fajl"].Value;
                this.path = matches[0].Groups["folder"].Value;
            }

            
        }
    }
}