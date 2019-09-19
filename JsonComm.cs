using System.Net;
using System.IO;
using System.Net.Security;
using System.Configuration;

namespace CoolTool
{
	static class JsonComm
	{
		private static string restURL = ConfigurationManager.AppSettings["memoQResourcesAPIURL"].ToString();
		public static string lastLocation = "";

		public static string SendRequest(byte[] data, bool isPost, string URLextension)
		{
			HttpWebRequest request = (HttpWebRequest)WebRequest.Create(restURL + URLextension);
			ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11;
			ServicePointManager.ServerCertificateValidationCallback = new
				RemoteCertificateValidationCallback(delegate { return true; });
			
			if (isPost)
			{
				request.Method = "POST";
				request.ContentType = "application/json; charset=utf-8";
				request.Accept = "application/json";
				if (data != null)
				{
					request.ContentLength = data.Length;
					using (var stream = request.GetRequestStream())
					{
						stream.Write(data, 0, data.Length);
					}
				}
				else
				{
					request.ContentLength = 0;
				}
			}

			HttpWebResponse response = (HttpWebResponse)request.GetResponse();
			GetLastLocationFromResponse(response);
			string responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
			return responseString;
		}

		private static void GetLastLocationFromResponse(HttpWebResponse response)
		{
			lastLocation = "";
			for (int i = 0; i < response.Headers.Count; ++i)
			{
				if (response.Headers.Keys[i].ToLower() == "location")
				{
					lastLocation = response.Headers[i];
					break;
				}
			}
		}

	}

}