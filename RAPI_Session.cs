using System;
using System.Collections.Generic;
using System.Text;
using System.Web.Script.Serialization;
using ResourcesAPI_TBList;
using ResourcesAPI_TMList;
using ResourcesAPI_TBEntry;
using System.Configuration;

namespace CoolTool
{
	static class RAPI_Session
	{

		public static string AccessToken = "";
		private static string testTBID = "";
		private static DateTime lastTest;
		public static string lastLocation = "";

		private static void Login()
		{
			LoginRequest loginRequest = new LoginRequest();
			loginRequest.username = ConfigurationManager.AppSettings["userLogin"].ToString();
			loginRequest.password = ConfigurationManager.AppSettings["userPassword"].ToString();


			string postData = new JavaScriptSerializer().Serialize(loginRequest);

			bool loginSuccess = false;
			int loginCounter = 0;

			do
			{
				loginCounter++;
				try
				{
					string response = JsonComm.SendRequest(Encoding.ASCII.GetBytes(postData), true, "/auth/login");
					LoginResponse loginResponse = new JavaScriptSerializer().Deserialize<LoginResponse>(response);
					AccessToken = loginResponse.AccessToken;

					response = JsonComm.SendRequest(null, false, "/tbs?authToken=" + AccessToken);
					TBListResponse[] tbListResponse = new JavaScriptSerializer().Deserialize<TBListResponse[]>(response);
					lastTest = DateTime.Now;

					if (tbListResponse.Length != 0)
					{
						testTBID = tbListResponse[0].TBGuid;
					}

					loginSuccess = true;
				}
				catch (Exception ex)
				{
					if (loginCounter < 10)
					{
						Log.AddLog("Login problem encountered. Retrying. " + ex.Message);
						System.Threading.Thread.Sleep(loginCounter * 200);
						loginSuccess = false;
					}
					else
					{
						throw new Exception("Login unsuccesful after 10 tries: " + ex.Message);
					}
				}
			}
			while (!loginSuccess);

		}

		private static void CheckLogin()
		{
			if (AccessToken == "" || testTBID == "")
			{
				Login();
			}

			if (testTBID == "")
			{
				throw new Exception("At least one TB needs to exist on the memoQ server.");
			}

			if (DateTime.Now - lastTest > new TimeSpan(0, 1, 0))
			{
				try
				{
					JsonComm.SendRequest(null, false, "/tbs/" + testTBID + "?authToken=" + AccessToken);
					lastTest = DateTime.Now;
				}
				catch (Exception ex)
				{
					if (ex.Message == "The remote server returned an error: (401) Unauthorized.")
					{
						Login();
						try
						{
							JsonComm.SendRequest(null, false, "/tbs/" + testTBID + "?authToken=" + AccessToken);
							lastTest = DateTime.Now;
						}
						catch (Exception exs)
						{
							throw new Exception("Login cannot be verified: " + exs.Message);
						}

					}
					else
					{
						throw new Exception("Login cannot be verified: " + ex.Message);
					}
				}
			}
		}

		public static void Dispose()
		{
			try
			{
				string response = JsonComm.SendRequest(null, true, "/auth/logout?authToken=" + AccessToken);
				AccessToken = "";
			}
			catch (Exception ex)
			{
				Log.AddLog("Logout failed: " + ex.Message, true);
			}
		}


		internal static bool UpdateTBEntry(string tbID, int entryID, TBEntry updateEntry)
		{
			CheckLogin();
			bool result = true;

			try
			{
				string postData = new JavaScriptSerializer().Serialize(updateEntry);
				string response = JsonComm.SendRequest(Encoding.UTF8.GetBytes(postData), true,
					"/tbs/" + tbID + "/entries/" + entryID + "/update?authToken=" + AccessToken);
				lastTest = DateTime.Now;
			}
			catch (Exception ex)
			{
				Log.AddLog("Entry update error: " + entryID + " - " + ex.Message, true);
				result = false;
			}
			return result;
		}

		internal static bool getTBEntry(string tbID, int entryID, out TBEntry tbEntry)
		{
			CheckLogin();
			bool result = true;
			tbEntry = new TBEntry();
			string response;

			try
			{
				string url = "/tbs/" + tbID + "/entries/" + entryID.ToString() + "?authToken=" + AccessToken;
				response = JsonComm.SendRequest(null, false, url);
				lastTest = DateTime.Now;
				tbEntry = new JavaScriptSerializer().Deserialize<TBEntry>(response);
			}
			catch (Exception ex)
			{
				if (ex.Message != "The remote server returned an error: (404) Not Found.")
				{
					//Log.AddLog("Entry retrieval error: " + entryID + " - " + ex.Message, true);
				}
				result = false;
			}
			return result;

		}

		internal static int GetLastEntryID(string tbID, out List<string> tbLanguages, out int tbSize)
		{
			CheckLogin();
			string response;
			string postData;
			tbLanguages = new List<string>();
			tbSize = 0;

			response = JsonComm.SendRequest(null, false, "/tbs/" + tbID + "?authToken=" + AccessToken);
			TBListResponse tblr = new JavaScriptSerializer().Deserialize<TBListResponse>(response);

			if (tblr.Languages.Count < 2) return 0;

			tbLanguages.AddRange(tblr.Languages);
			tbSize = tblr.NumEntries;
			string lang = tblr.Languages[0];

			ResourcesAPI_TBEntry.TermItem tbeTerm = new ResourcesAPI_TBEntry.TermItem();
			tbeTerm.Text = "Naezttaláldkitenagyokos";
			ResourcesAPI_TBEntry.Language tbeLang = new ResourcesAPI_TBEntry.Language();
			tbeLang.language = lang;
			tbeLang.TermItems = new List<ResourcesAPI_TBEntry.TermItem>();
			tbeLang.TermItems.Add(tbeTerm);

			TBEntry te = new TBEntry();
			te.Created = DateTime.Now.ToString();
			te.Creator = ConfigurationManager.AppSettings["userLogin"].ToString();
			te.Languages = new List<ResourcesAPI_TBEntry.Language>();
			te.Languages.Add(tbeLang);

			postData = new JavaScriptSerializer().Serialize(te);
			int entryID = 0;

			try
			{
				response = JsonComm.SendRequest(Encoding.UTF8.GetBytes(postData), true, "/tbs/" + tbID + "/entries/create?authToken=" + AccessToken);
				string lastLocation = JsonComm.lastLocation;
				entryID = Convert.ToInt32(lastLocation.Substring(lastLocation.LastIndexOf('/') + 1));
				response = JsonComm.SendRequest(null, true, "/tbs/" + tbID + "/entries/" + entryID.ToString() + "/delete?authToken=" + AccessToken);
				lastTest = DateTime.Now;
			}
			catch (Exception ex)
			{
				if (ex.Message == "The remote server returned an error: (404) Not Found.")
				{
					Log.AddLog("The term base is probably open for editing in memoQ. Close it first.", true);
				}
				throw new Exception(ex.Message);
			}

			return entryID - 1;

			//TBLookupRequest tbLookupRequest = new TBLookupRequest();
			//tbLookupRequest.SourceLanguage = lang;
			//tbLookupRequest.TargetLanguage = tblr.Languages[1];
			//string segment = "<seg>Naezttaláldkitenagyokos</seg>";
			//tbLookupRequest.Segments.Add(segment);
			//postData = new JavaScriptSerializer().Serialize(tbLookupRequest);
			//response = JsonComm.SendRequest(Encoding.UTF8.GetBytes(postData), true, "/tbs/" + tbID + "/lookupterms?authToken=" + AccessToken);

			//TBLookupResponse tblur = new JavaScriptSerializer().Deserialize<TBLookupResponse>(response);

			//var hits = tblur.Result[0].TBHits[0];


			//foreach (var hit in hits)
			//{
			//    entryID = hit.Entry.Id;
			//    try
			//    {

			//        lastTest = DateTime.Now;
			//    }
			//    catch (Exception ex)
			//    {
			//        Log.AddLog("The last term checker expression could not be deleted. " + ex.Message, true);
			//    }
			//}
		}



		internal static TMListResponse[] generateTMlist()
		{
			CheckLogin();
			//Log.AddLog("TM list is to be requested.", false);
			string response = JsonComm.SendRequest(null, false, "/tms?authToken=" + AccessToken);
			lastTest = DateTime.Now;
			//Log.AddLog("TM list requested successfuly.", false);
			TMListResponse[] tmListResponse = new JavaScriptSerializer().Deserialize<TMListResponse[]>(response);
			return tmListResponse;
		}

		internal static TBListResponse[] generateTBlist()
		{
			CheckLogin();
			string response = JsonComm.SendRequest(null, false, "/tbs?authToken=" + AccessToken);
			TBListResponse[] tbListResponse = new JavaScriptSerializer().Deserialize<TBListResponse[]>(response);
			lastTest = DateTime.Now;

			return tbListResponse;
		}


	}

}
