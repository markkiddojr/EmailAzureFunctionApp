using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Threading;

namespace EmailApp
{
    public static class SheetsFunction
    {
        #region Declare class variables (Don't forget to update TODO section! - check View > Task List)
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };

        ///TODO: Replace with correct information///////////////////////////////////////////////////////////////////////////////////////////////////////////
        static string ApplicationName = Environment.GetEnvironmentVariable("ApplicationName").ToString();
        static string sheetId = Environment.GetEnvironmentVariable("SheetId").ToString();
        static string clientId = Environment.GetEnvironmentVariable("ClientId").ToString();         // From https://console.developers.google.com  
        static string clientSecret = Environment.GetEnvironmentVariable("ClientSecret").ToString();
        static string range = Environment.GetEnvironmentVariable("Range").ToString();               //"Class Data!A2:E";          
        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        static UserCredential credential;
        #endregion

        #region Define Azure Function. Rout will look as follows: {domain}.com/{function endpoint}?email={user@email.com}&answer={yes or no}
        [FunctionName("SheetsFunction")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            string email = req.Query["email"].ToString().ToLower();
            string answer = req.Query["answer"].ToString().ToLower() == "yes" ? "Yes" : "No";

            //some of the following code provided by Google APIs https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values/append
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = GetCredential(),
                ApplicationName = ApplicationName,
            });
            
            //Append our email and answer to the request body of values to add
            ValueRange requestBody = new ValueRange();
            var values = new[] { new { email, answer } };
            requestBody.Values.Add(values);

            //Create our request and set up the input options
            SpreadsheetsResource.ValuesResource.AppendRequest request = service.Spreadsheets.Values.Append(requestBody, sheetId, range);
            SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum valueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;  
            SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum insertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS; 
            request.ValueInputOption = valueInputOption;
            request.InsertDataOption = insertDataOption;

            //Obtain our response
            AppendValuesResponse response = await request.ExecuteAsync();

            //If everything is good we should have a valid response
            return response != null
                ? (ActionResult)new OkObjectResult(response)
                : new BadRequestObjectResult("Something went wrong.");
        }
        #endregion

        #region Helper Methods
        private static UserCredential GetCredential()
        {
            var task = GoogleWebAuthorizationBroker.AuthorizeAsync(
                new ClientSecrets
                {
                    ClientId = clientId,
                    ClientSecret = clientSecret
                }, 
                Scopes, 
                Environment.UserName, 
                CancellationToken.None, 
                new FileDataStore("MyAppsToken")
            );

            task.Wait();
            credential = task.Result;
            return credential;
        }
        #endregion
    }
}