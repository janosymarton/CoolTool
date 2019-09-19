using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace CoolTool
{
    public class Language
    {
        public string DisplayName { get; set; }
        public string ISOCode { get; set; }
        public string MemoQCode { get; set; }
        public int LCID { get; set; }
    }
    
    public static class LanguageHelper
    {

        private const string LanguagesFileName = "languages.csv";
        public static List<Language> Languages;
        public static bool Initialize()
        {
            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                string resourceName = "CoolTool.languages.csv";
                string[] assemblies = assembly.GetManifestResourceNames();
                

                Languages = new List<Language>();
                
                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                using (TextReader rd = new StreamReader(stream, Encoding.UTF8))
                {
                    string line = "";
                    while ((line = rd.ReadLine()) != null)
                    {
                        var d = line.Split(';');
                        if (d.Length > 0)
                        {
                            int lcid = 0;
                            if (d[3] != "")
                                lcid = int.Parse(d[3]);
                            Language newlang = new Language()
                                {
                                    DisplayName = d[0],
                                    ISOCode = d[1],
                                    MemoQCode = d[2],
                                    LCID = lcid,
                                };
                            Languages.Add(newlang);
                        }
                    }
                }
                Log.AddLog("Language list initialization completed.", false);
                return true;
            }
            catch (Exception ex)
            {
                Log.AddLog("Language list initialization error." + ex.Message, true);
                return false;
            }
        }

        public static Language GetLanguageByMemoQCode(string memoQCode)
        {
            foreach (Language language in Languages)
            {
                if (language.MemoQCode.ToLower() == memoQCode.ToLower())
                    return language;
            }
            Language unkown = new Language();
            unkown.DisplayName = "Unknown";
            return unkown;
        }

        public static Language GetLanguageByISOCode(string isoCode)
        {
            foreach (Language language in Languages)
            {
                if (language.ISOCode.ToLower() == isoCode.ToLower())
                    return language;
            }
            return null;
        }

        public static Language GetLanguageByDisplayName(string DisplayName)
        {
            foreach (Language language in Languages)
            {
                if (language.DisplayName.ToLower() == DisplayName.ToLower())
                    return language;
            }
            return null;
        }

        internal static string[] LanguageNames()
        {
            List<string> languageNames = new List<string>();

            foreach (Language language in Languages)
            {
                languageNames.Add(language.DisplayName);
            }
            return languageNames.ToArray();
        }

        internal static Language[] GetLanguageList(CheckedListBox.CheckedItemCollection checkedItemCollection)
        {
            List<Language> languages = new List<Language>();
            foreach (object itemChecked in checkedItemCollection)
            {
                languages.Add(LanguageHelper.GetLanguageByDisplayName(itemChecked.ToString()));
            }
            return languages.ToArray();
        }
    }
}
