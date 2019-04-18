using System;
using System.Text;
using System.IO;
using UnityEngine;
using UnityEditor;
using Google.GData.Client; //Google Sheets Api v3(legacy)
using Google.GData.Spreadsheets;
//for google sheets api v3, see https://developers.google.com/sheets/api/v3/authorize
using UnityEngine.Networking;
using System.Collections;

//process flow
//getAuthorizedService() -> requestSheets()

namespace SNUPlugin
{
    public class SNUPlugin
    {
        //private const string AUTH = "credentials.json";
        private const string AUTH = @"E:\VSProjects\SNUPlugin\credentials.json";
        //read credentials from json
        private Credentials getAuths(string path)
        {
            if (!File.Exists(path))
            {
                EditorUtility.DisplayDialog("SNUPlugin", "구글 계정 접근 권한을 위한 credentials.json을 열 수 없습니다.\nhttps://developers.google.com/sheets/api/quickstart/dotnet을 참조하세요.", "닫기");
                return null;
            }
            FileStream f = new FileStream(path, FileMode.Open);
            StreamReader sr;
            if (!f.CanRead) return null;
            sr = new StreamReader(f, Encoding.Default);
            try
            {
                string str = sr.ReadToEnd();
                sr.Close();
                f.Close();
                f.Dispose();
                return JsonUtility.FromJson<Credentials>(str);
            } catch
            {
                sr.Close();
                f.Close();
                f.Dispose();
                return null;
            }
        }
        //get authorization for google
        private SpreadsheetsService getAuthorizedService()
        {
            //Get Auths
            Credentials cred = getAuths(AUTH);
            if (cred == null) return null;

            ////////////////////////////////////////////////////////////////////////////
            // STEP 1: Configure how to perform OAuth 2.0
            ////////////////////////////////////////////////////////////////////////////

            // TODO: Update the following information with that obtained from
            // https://code.google.com/apis/console. After registering
            // your application, these will be provided for you.

            string CLIENT_ID = cred.installed.client_id;

            // This is the OAuth 2.0 Client Secret retrieved
            // above.  Be sure to store this value securely.  Leaking this
            // value would enable others to act on behalf of your application!
            string CLIENT_SECRET = cred.installed.client_secret;

            // Space separated list of scopes for which to request access.
            string SCOPE = "https://spreadsheets.google.com/feeds https://docs.google.com/feeds";

            // This is the Redirect URI for installed applications.
            // If you are building a web application, you have to set your
            // Redirect URI at https://code.google.com/apis/console.
            string REDIRECT_URI = "urn:ietf:wg:oauth:2.0:oob";

            ////////////////////////////////////////////////////////////////////////////
            // STEP 2: Set up the OAuth 2.0 object
            ////////////////////////////////////////////////////////////////////////////

            // OAuth2Parameters holds all the parameters related to OAuth 2.0.
            OAuth2Parameters parameters = new OAuth2Parameters();

            // Set your OAuth 2.0 Client Id (which you can register at
            // https://code.google.com/apis/console).
            parameters.ClientId = CLIENT_ID;

            // Set your OAuth 2.0 Client Secret, which can be obtained at
            // https://code.google.com/apis/console.
            parameters.ClientSecret = CLIENT_SECRET;

            // Set your Redirect URI, which can be registered at
            // https://code.google.com/apis/console.
            parameters.RedirectUri = REDIRECT_URI;

            ////////////////////////////////////////////////////////////////////////////
            // STEP 3: Get the Authorization URL
            ////////////////////////////////////////////////////////////////////////////

            // Set the scope for this particular service.
            parameters.Scope = SCOPE;

            // Get the authorization url.  The user of your application must visit
            // this url in order to authorize with Google.  If you are building a
            // browser-based application, you can redirect the user to the authorization
            // url.
            string authorizationUrl = OAuthUtil.CreateOAuth2AuthorizationUrl(parameters);
            string msg = string.Empty;
            msg = "OAuth-request 토큰을 부여받기 위하여 이곳을 클릭하시어\n"
                 + "복사한 값을 아래에 입력해주세요."; ;
            parameters.AccessCode = AuthForm.getToken(authorizationUrl, msg);

            ////////////////////////////////////////////////////////////////////////////
            // STEP 4: Get the Access Token
            ////////////////////////////////////////////////////////////////////////////

            // Once the user authorizes with Google, the request token can be exchanged
            // for a long-lived access token.  If you are building a browser-based
            // application, you should parse the incoming request token from the url and
            // set it in OAuthParameters before calling GetAccessToken().
            OAuthUtil.GetAccessToken(parameters);
            string accessToken = parameters.AccessToken;
            Debug.Log("OAuth Access Token: " + accessToken);

            ////////////////////////////////////////////////////////////////////////////
            // STEP 5: Make an OAuth authorized request to Google
            ////////////////////////////////////////////////////////////////////////////

            // Initialize the variables needed to make the request
            GOAuth2RequestFactory requestFactory =
                new GOAuth2RequestFactory(null, "MySpreadsheetIntegration-v1", parameters);
            SpreadsheetsService service = new SpreadsheetsService("MySpreadsheetIntegration-v1");
            service.RequestFactory = requestFactory;

            // Make the request to Google
            // See other portions of this guide for code to put here...
            return service;
        }
        //request spreadsheets data
        public bool requestSheets()
        {
            SpreadsheetsService service = getAuthorizedService();

            if (service == null) return false;
            // TODO: Authorize the service object for a specific user (see other sections)

            // Instantiate a SpreadsheetQuery object to retrieve spreadsheets.
            SpreadsheetQuery query = new SpreadsheetQuery();

            // Make a request to the API and get all spreadsheets.
            SpreadsheetFeed feed = service.Query(query);

            if (feed.Entries.Count == 0)
            {
                // TODO: There were no spreadsheets, act accordingly.
            }

            // TODO: Choose a spreadsheet more intelligently based on your
            // app's needs.
            SpreadsheetEntry spreadsheet = (SpreadsheetEntry)feed.Entries[0];
            Debug.Log(spreadsheet.Title.Text);

            // Get the first worksheet of the first spreadsheet.
            // TODO: Choose a worksheet more intelligently based on your
            // app's needs.
            WorksheetFeed wsFeed = spreadsheet.Worksheets;
            WorksheetEntry worksheet = (WorksheetEntry)wsFeed.Entries[0];

            // Define the URL to request the list feed of the worksheet.
            AtomLink listFeedLink = worksheet.Links.FindService(GDataSpreadsheetsNameTable.ListRel, null);

            // Fetch the list feed of the worksheet.
            ListQuery listQuery = new ListQuery(listFeedLink.HRef.ToString());
            ListFeed listFeed = service.Query(listQuery);

            // Iterate through each row, printing its cell values.
            StringBuilder stbRes = new StringBuilder();
            foreach (ListEntry row in listFeed.Entries)
            {
                // Print the first column's cell value
                stbRes.AppendLine(row.Title.Text);
                // Iterate over the remaining columns, and print each cell value
                foreach (ListEntry.Custom element in row.Elements)
                {
                    stbRes.AppendLine(element.Value);
                }
            }
            EditorUtility.DisplayDialog("SNUPlugin", stbRes.ToString(), "OK");
            return true;
        }
    }

    [System.Serializable]
    class Credentials
    {
        public credDetails installed;
        [System.Serializable]
        public class credDetails
        {
            public string client_id;
            public string project_id;
            public string auth_uri;
            public string token_uri;
            public string auth_provider_x509_cert_url;
            public string client_secret;
            public string[] redirect_uris;
        }
    }
}
