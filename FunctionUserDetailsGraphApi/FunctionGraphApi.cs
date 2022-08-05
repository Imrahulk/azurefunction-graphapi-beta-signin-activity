using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Net.Http.Headers;
using System.Net.Http;
using FunctionUserDetailsGraphApi.Models;
using System.Linq;
using System.Data;
using System.Reflection;
using ClosedXML.Excel;
using System.IO;
using Azure.Storage.Blobs;

namespace FunctionUserDetailsGraphApi
{

    public static class FunctionGraphApi
    {
        private static HttpClient httpClient = new HttpClient();
        private static AuthenticationContext context = null;
        private static ClientCredential credential = null;


        [FunctionName("FunctionGraphApi")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            context = new AuthenticationContext(Constants.Authority);
            credential = new ClientCredential(Constants.ClientId, Constants.ClientSecret);
            Task<string> token = GetAuthToken();
            token.Wait();

            var userDetails = await GetUserDetails(token.Result, Constants.BetaBaseUrl, new List<Value>());
            var flaggedUsers = userDetails.Where(x => x.signInActivity != null && x.manager != null).Where(x => x.signInActivity.lastSignInDateTime < Constants.LastLoginDate).ToList();

            var signInReport = GenerateSignInReportList(flaggedUsers);

            DataTable dtUserSignInReport = ConvertListToDataTable(signInReport);
            UploadExcelReportToBlobStorage(log, dtUserSignInReport);

            var managerDetails = GenerateDistinctManagerRecords(signInReport);

            return new OkObjectResult(managerDetails);
        }

        private static async Task<string> GetAuthToken()
        {
            AuthenticationResult result = await context.AcquireTokenAsync(Constants.Resource, credential);
            return result.AccessToken;
        }
        private static async Task<List<Value>> GetUserDetails(string result, string url, List<Value> recordCollection)
        {

            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", result);
            var data = await httpClient.GetAsync(url);
            if (data.Content != null)
            {
                string users = await data.Content.ReadAsStringAsync();
                UserDetails userDetails = JsonConvert.DeserializeObject<UserDetails>(users);
                if (userDetails.value.Count > 0)
                {
                    foreach (var item in userDetails.value)
                        recordCollection.Add(item);
                }
                if (userDetails.OdataNextLink != null)
                    await GetUserDetails(result, userDetails.OdataNextLink, recordCollection);
            }
            return recordCollection;
        }

        private static DataTable ConvertListToDataTable<T>(List<T> userDetails)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
                dataTable.Columns.Add(prop.Name);

            foreach (T item in userDetails)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                    values[i] = Props[i].GetValue(item, null);

                dataTable.Rows.Add(values);
            }
            return dataTable;

        }
        private static List<SignInReport> GenerateSignInReportList(List<Value> userDetails)
        {
            List<SignInReport> signInReport = new List<SignInReport>();

            if (userDetails != null)
            {
                if (userDetails.Count != 0)
                {
                    foreach (var item in userDetails)
                    {
                        SignInReport signInDetails = new SignInReport();
                        signInDetails.DisplayName = item.displayName;
                        signInDetails.Id = item.id;
                        signInDetails.LastSignInDateTime = item.signInActivity.lastSignInDateTime;
                        if (item.manager != null)
                        {
                            signInDetails.ManagerName = item.manager.givenName;
                            signInDetails.ManagerEmail = item.manager.mail;
                        }
                        signInReport.Add(signInDetails);
                    }

                }

            }
            return signInReport.OrderBy(x => x.LastSignInDateTime).ToList();
        }

        private static void UploadExcelReportToBlobStorage(ILogger log, DataTable datatable)
        {
            try
            {
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    //Generate Excel
                    var workbook = new XLWorkbook();
                    workbook.Worksheets.Add(datatable);
                    workbook.Worksheets.First().Columns().AdjustToContents();
                    workbook.SaveAs(memoryStream);

                    //Upload excel to blob storage
                    BlobServiceClient serviceClient = new BlobServiceClient(Constants.ConnectionString);
                    BlobContainerClient containerClient = serviceClient.GetBlobContainerClient(Constants.ContainerName);
                    BlobClient blobClient = containerClient.GetBlobClient(Constants.BlobName);

                    bool isExist = blobClient.Exists();
                    if (isExist)
                        blobClient.DeleteIfExists();

                    memoryStream.Position = 0;
                    blobClient.Upload(memoryStream);
                }
            }
            catch (Exception ex)
            {
                log.LogInformation(ex.ToString());

            }
        }

        private static List<ManagerDetailsResponse>  GenerateDistinctManagerRecords(List<SignInReport> signInReport)
        {
            IEnumerable<SignInReport> distinctManagerDetails = signInReport.DistinctBy(x => x.ManagerEmail).Distinct().ToList();
            List<ManagerDetailsResponse> managerResponse = new List<ManagerDetailsResponse>();

            foreach (var item in distinctManagerDetails)
            {
                ManagerDetailsResponse manager = new ManagerDetailsResponse();
                if (item.ManagerName != null && item.ManagerEmail != null)
                {
                    manager.ManagerName = item.ManagerName;
                    manager.ManagerEmail = item.ManagerEmail;
                    managerResponse.Add(manager);
                }
            }
            return managerResponse;

        }

    }

}
