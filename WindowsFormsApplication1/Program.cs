using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SheetsQuickstart
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "Google Sheets API .NET Quickstart";

        static void Main(string[] args)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("client_secret_424891321849-fo9l06m0jeke5hsvgtuisf97fnd4inf6.apps.googleusercontent.com.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            //https://docs.google.com/a/theryanzoo.com/spreadsheets/d/1fwoB2S-j6k10TYA-5wXtmVI2NXpP2iYPBZYSHCWNO-M/edit?usp=sharing
            // Define request parameters.
            String spreadsheetId = "1fwoB2S-j6k10TYA-5wXtmVI2NXpP2iYPBZYSHCWNO-M";
            String range = "A1:b3";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);


//sheet_metadata = service.spreadsheets().get(spreadsheetId = spreadsheet_id).execute()
//sheets = sheet_metadata.get('sheets', '')
//title = sheets[0].get("properties", { }).get("title", "Sheet1")
//sheet_id = sheets[0].get("properties", { }).get("sheetId", 0)
            SpreadsheetsResource.GetRequest xxx=service.Spreadsheets.Get(spreadsheetId);



            // https://docs.google.com/spreadsheets/d/1fwoB2S-j6k10TYA-5wXtmVI2NXpP2iYPBZYSHCWNO-M/edit
            ValueRange response = request.Execute();




            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                Console.WriteLine("Name, Major");
                foreach (var row in values)
                {
                    // Print columns A and E, which correspond to indices 0 and 4.
                    Console.WriteLine("{0}, {1}", row[0], row[1]);
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }
            Console.Read();


        }
    }
}