using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace SheetsQuickstart
{
    class Program
    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "GoogleSheetsWeeklyReporter";

        static void Main(string[] args)
        {
            try
            {
                UserCredential credential;
                
                using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                {
                    string credPath = "token.json";
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.FromStream(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                    Console.WriteLine("Credential file saved to: " + credPath);
                }

                var service = new SheetsService(new BaseClientService.Initializer
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName
                });

                //String spreadsheetId = "1wcXSo-0LVGaNowKnc-geEC81TPEzpBHpS3odPP1X4qg";
                //String range1 = "List!A2:C";
                //SpreadsheetsResource.ValuesResource.GetRequest request =
                //    service.Spreadsheets.Values.Get(spreadsheetId, range1);

                //ValueRange response = request.Execute();
                //IList<IList<Object>> values = response.Values;
                //if (values == null || values.Count == 0)
                //{
                //    Console.WriteLine("No data found.");
                //    return;
                //}
                //Console.WriteLine("Name, Major");
                //foreach (var row in values)
                //{
                //    Console.WriteLine("{0}, {1}", row[1], row[2]);
                //}

                String spreadsheetId2 = "1wcXSo-0LVGaNowKnc-geEC81TPEzpBHpS3odPP1X4qg";
                String range2 = "List2!F5";  // update cell F5 
                ValueRange valueRange = new ValueRange();
                valueRange.MajorDimension = "COLUMNS";//"ROWS";//COLUMNS

                var oblist = new List<object>() { "My Cell Text" };
                valueRange.Values = new List<IList<object>> { oblist };

                SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId2, range2);
                update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                UpdateValuesResponse result2 = update.Execute();

                Console.WriteLine("done!");
            }
            catch (FileNotFoundException ex) { Console.WriteLine(ex.Message);
            }
        }
    }
}
