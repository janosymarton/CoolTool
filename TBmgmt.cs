using System;
using ResourcesAPI_TBEntry;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace CoolTool
{
    internal class TBmgmt
    {
        private Guid tbID;
        private string substText;
        private string searchText;
        private int caseSens;
        private int matching;
        private List<string> tbLanguages;
        private Dictionary<string, int> missingTerms = new Dictionary<string, int>();

        public TBmgmt(string substText, string searchText, Guid tbID, int caseSens, int matching)
        {
            this.substText = substText;
            this.searchText = searchText;
            this.tbID = tbID;
            this.matching = matching;
            this.caseSens = caseSens;

            ProcessTB();
            Program.mainWindow.updateProgress(100);
        }
    
		private void ProcessTB()
		{
			int processedEntries = 0;
			int failedEntries = 0;
			int tbSize;
			Dictionary<int, TBEntry> tbEntriesToUpdate = new Dictionary<int, TBEntry>();

			int lastID = RAPI_Session.GetLastEntryID(tbID.ToString(), out tbLanguages, out tbSize);
            lastID = RAPI_Session.GetLastEntryID(tbID.ToString(), out tbLanguages, out tbSize); // need to run two queries due to server bug
            if (lastID > 0)
			{
				foreach (string lang in tbLanguages)
				{
					missingTerms.Add(lang, 0);
				}

				Program.mainWindow.updateProgress(10);              

                for (int i = 0; i < lastID; i++)
				{
					try
					{
						TBEntry tbEntry;
						if (RAPI_Session.getTBEntry(tbID.ToString(), i, out tbEntry))
						{
							processedEntries++;
							if (CheckEntry(ref tbEntry))
							{
								tbEntriesToUpdate.Add(i, tbEntry);
                                Log.AddLog("TBEntry fixed: " + i + ". Total: " + lastID);
							}

							Program.mainWindow.updateProgress(Convert.ToInt32(10 + 90 * processedEntries / tbSize));
						}
					}
					catch (Exception ex)
					{
						Log.AddLog("Error while processing TB entry " + i.ToString() + " - " + ex.Message, true);
						failedEntries++;
					}
				}
			}

			foreach (int entryID in tbEntriesToUpdate.Keys)
			{
				TBEntry tbe = tbEntriesToUpdate[entryID];
				if (!RAPI_Session.UpdateTBEntry(tbID.ToString(), entryID, tbe))
				{
					failedEntries++;
				}
			}

			Log.AddLog(processedEntries.ToString() + " termbase entries processed." + (failedEntries == 0 ? "" :
										(failedEntries.ToString() + " entries failed")), failedEntries != 0);

			bool isLangMessageDisplayed = false;
			foreach (string key in missingTerms.Keys)
			{
				if (missingTerms[key] > 0)
				{
					Log.AddLog("For " + LanguageHelper.GetLanguageByMemoQCode(key).DisplayName + " " + missingTerms[key] +
											" term" + (missingTerms[key] == 1 ? " was" : "s were") + " marked with " +
											substText + ".", false);
					isLangMessageDisplayed = true;
				}
			}
			if (!isLangMessageDisplayed)
				Log.AddLog("No missing terms.");
		}


		private bool CheckEntry(ref TBEntry tbEntry)
		{
			bool result = false;

			foreach (string tbLang in tbLanguages)
			{
				bool isLangInEntry = false;
				ResourcesAPI_TBEntry.Language entryLangMatching = new ResourcesAPI_TBEntry.Language();

				foreach (ResourcesAPI_TBEntry.Language entryLang in tbEntry.Languages)
				{
					if (entryLang.language == tbLang)
					{
						isLangInEntry = true;
						entryLangMatching = entryLang;
						break;
					}
				}

				if (isLangInEntry)
				{
					if (entryLangMatching.TermItems == null)
					{
						entryLangMatching.TermItems = new List<TermItem>();
					}

					if (entryLangMatching.TermItems.Count == 0)
					{
						TermItem tbeTerm = getSubstTermItem();
						entryLangMatching.TermItems.Add(tbeTerm);
						missingTerms[tbLang]++;
                        RAPI_Session.UpdateTBEntry(tbID.ToString(), tbEntry.Id, tbEntry);   // This line is to work around the bug that we cannot add multiple terms
                        RAPI_Session.getTBEntry(tbID.ToString(), tbEntry.Id, out tbEntry); // This line is to work around the bug that we cannot add multiple terms
                        result = true;
					}
					else if (entryLangMatching.TermItems.Count == 1)
					{
						if (entryLangMatching.TermItems[0].Text == searchText)
						{
							entryLangMatching.TermItems[0].Text = substText;
							missingTerms[tbLang]++;
							result = true;
						}
					}
				}
				else
				{
					ResourcesAPI_TBEntry.Language tbeLang = new ResourcesAPI_TBEntry.Language();
					tbeLang.language = tbLang;
					tbeLang.Id = -1;
					tbeLang.TermItems = new List<TermItem>();
                    TermItem tbeTerm = getSubstTermItem();
                    tbeLang.TermItems.Add(tbeTerm);
                    tbEntry.Languages.Add(tbeLang);
                    missingTerms[tbLang]++;
                    RAPI_Session.UpdateTBEntry(tbID.ToString(), tbEntry.Id, tbEntry);  // This line is to work around the bug that we cannot add multiple terms
                    RAPI_Session.getTBEntry(tbID.ToString(), tbEntry.Id, out tbEntry); // This line is to work around the bug that we cannot add multiple terms
                    
                    result = true;
                }
			}

			return result;
		}

		private TermItem getSubstTermItem()
		{
			TermItem tbeTerm = new TermItem();
			tbeTerm.Text = substText;
			tbeTerm.Id = -1;
			tbeTerm.CaseSense = caseSens;
			tbeTerm.PartialMatch = matching; 
			return tbeTerm;
		}
	}
}