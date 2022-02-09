using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Util.Store;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;

namespace SheetsPersist
{
	public static partial class GoogleSheets
	{
		static string ApplicationName = "Google Sheets Persist";
		
		static string[] Scopes = { SheetsService.Scope.Spreadsheets };

		static GoogleSheets()
		{
			InitializeService(GetUserCredentials());
		}

		private static void InitializeService(UserCredential credential)
		{
			service = new SheetsService(new BaseClientService.Initializer()
			{
				HttpClientInitializer = credential,
				ApplicationName = ApplicationName,
			});
		}

		private static UserCredential GetUserCredentials()
		{
			UserCredential credential;
			using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
			{
				// The file token.json stores the user's access and refresh tokens, and is created
				// automatically when the authorization flow completes for the first time.
				string credentialPath = "token.json";
				ClientSecrets secrets = GoogleClientSecrets.FromStream(stream).Secrets;
				FileDataStore dataStore = new FileDataStore(credentialPath, true);
				Debug.WriteLine("GoogleWebAuthorizationBroker.AuthorizeAsync...");
				credential = GoogleWebAuthorizationBroker.AuthorizeAsync(secrets, Scopes, "user", CancellationToken.None, dataStore).Result;
				Debug.WriteLine("Credential file saved to: " + credentialPath);
			}

			return credential;
		}
	}
}

