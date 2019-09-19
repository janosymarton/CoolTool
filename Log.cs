using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace CoolTool
{

    class Log
    {
        public static string logName = "CoolToolLog";
        public static string logPath = Path.Combine(Path.GetTempPath(), logName + ".txt");

        public static bool isLog = false;
        public static bool isProgress = false;

        public Log()
        {
               if(File.Exists(logPath))
               {
                   FileInfo fi = new FileInfo(logPath);
                   if(fi.Length > 100000)
                   {
                       File.Move(logPath, Path.Combine(Path.GetTempPath(), logName + DateTime.Now.ToShortDateString() + "_" + DateTime.Now.ToShortTimeString() + ".txt"));
                   }
                   
               }
        }


        public static void AddLog(string logMessage, bool type = false)
        {
            logMessage = DateTime.Now.ToString() + " - " + logMessage;

            if (isLog)
            {
                using (StreamWriter w = File.AppendText(logPath))
                {
                    w.WriteLine((type ? "!!!" : "") + logMessage);
                }
            }
            if (isProgress)
            {
                Program.mainWindow.AddMessage(logMessage + Environment.NewLine, type);
            }

        }


    }
}
