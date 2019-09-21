/*
 * Created By Pratik Khandelwal
 */
using Microsoft.Extensions.Logging;
using Newtonsoft.Json.Linq;
using PresentationService.Models;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web;

namespace PresentationService.Services
{
    public interface INetworkService
    {
        Task<List<ViewData>> GetSheet(RequestParams dashboardParams);
        Task<List<ViewData>> GetViewsForWorkbooks(List<string> workbookList);
        Task<List<string>> GetWorkbooksForSite(List<string> dashboardList);
        Task<List<ViewData>> QueryViewData(List<ViewData> viewList, RequestParams globalFilters);
    }

    public class NetworkService : INetworkService
    {
        // Static Properties
        private static string urlString = @"";
        private static System.Net.Http.HttpClient httpClient = new System.Net.Http.HttpClient();
        private static string siteId { get; set; }
        private static string authToken { get; set; }

        private readonly ILogger _logger;
        public NetworkService(ILogger<NetworkService> logger)
        {
            _logger = logger;
        }

        static NetworkService()
        {
            httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        }

        public async Task<List<ViewData>> GetSheet(RequestParams dashboardParams)
        {
            List<string> workbookList = await GetWorkbooksForSite(dashboardParams.dashboardList);
            List<ViewData> viewList = await GetViewsForWorkbooks(workbookList);
            List<ViewData> viewStreamList = await QueryViewData(viewList, dashboardParams);
            return viewStreamList;
        }

        //To-do: Try to optimize this with yield
        public async Task<List<ViewData>> QueryViewData(List<ViewData> viewList, RequestParams globalFilters)
        {
            string viewUrl = "";
            
            viewUrl = urlString + @"views/{1}/data?" + globalFilters.globalFilters;
            foreach (var view in viewList)
            {
                string url = String.Format(viewUrl, siteId, view.ViewId);
                try
                {
                    HttpResponseMessage httpResponse = await httpClient.GetAsync(new Uri(url));
                    HttpContent httpContent = httpResponse.Content;
                    Stream httpContentStream = await httpContent.ReadAsStreamAsync();
                    view.DataStream = httpContentStream;
                }
                catch (Exception e)
                {
                    _logger.LogError("In QueryViewData");
                    _logger.LogError(e.Message);
                    throw e;
                }
            }
            return viewList;
        }

        public async Task<List<ViewData>> GetViewsForWorkbooks(List<string> workbookList)
        {
            List<ViewData> viewList = new List<ViewData>();
            foreach (string workbookId in workbookList)
            {
                string url = String.Format(urlString + @"workbooks/{1}/views", siteId, workbookId);
                try
                {
                    HttpResponseMessage httpResponse = await httpClient.GetAsync(new Uri(url));
                    string httpResponseMessage = await httpResponse.Content.ReadAsStringAsync();
                    JObject responseBody = JObject.Parse(httpResponseMessage);
                    JArray viewArray = (JArray)responseBody["views"]["view"];
                    foreach (var item in viewArray)
                    {
                        JArray tagArray = (JArray)item["tags"]["tag"];
                        try
                        {
                            string chartType = ((String)tagArray[0]["label"]).Split(":")[1];
                            string primaryAxis = ((string)tagArray[1]["label"]).Split(":")[1];
                            List<string> seriesList = ((string)tagArray[2]["label"]).Split(":")[1].Split("-").ToList();
                            viewList.Add(new ViewData((string)item["name"], (string)item["id"], chartType, primaryAxis, seriesList));
                        }
                        catch (Exception e)
                        {
                            _logger.LogError("Error in view Tags. ViewId:" + (string)item["name"]);
                            viewList.Add(new ViewData((string)item["name"], (string)item["id"], "", "", null));
                        }

                    }
                }
                catch (Exception e)
                {
                    _logger.LogError("Error while getting views for workbook " + workbookId);
                    _logger.LogError(e.Message);
                    continue;
                }
            }
            return viewList;
        }

        public async Task<List<string>> GetWorkbooksForSite(List<string> dashboardList)
        {
            try
            {
                if (siteId == null)
                {
                    await TableauAuthenticate();
                }
                String url = String.Format(urlString + @"workbooks/", siteId);
                HttpResponseMessage httpResponse = await httpClient.GetAsync(new Uri(url));
                string httpResponseMessage = await httpResponse.Content.ReadAsStringAsync();
                JObject responseBody = JObject.Parse(httpResponseMessage);
                if (httpResponseMessage.Contains("Unauthorized Access"))
                {
                    await TableauAuthenticate();
                    url = String.Format(urlString + @"workbooks/", siteId);
                    httpResponse = await httpClient.GetAsync(new Uri(url));
                    httpResponseMessage = await httpResponse.Content.ReadAsStringAsync();
                    responseBody = JObject.Parse(httpResponseMessage);
                }
                JArray workbookArray = (JArray)responseBody["workbooks"]["workbook"];
                List<string> workbookList = workbookArray.Where(c => dashboardList.Contains((string)c["contentUrl"])).Select(c => (string)c["id"]).ToList();
                return workbookList;
            }
            catch (Exception e)
            {
                _logger.LogError("Error while getting workbooks for site");
                _logger.LogError(e.Message);
                throw e;
            }
        }

        private async Task TableauAuthenticate()
        {
            HttpResponseMessage httpResponse = new HttpResponseMessage();
            //Crate a Json Payload string to send your credentials

    //        {
    //            "credentials": {
    //                "name": "admin",
    //                "password": "p@ssword",
    //                "site": {
    //                    "contentUrl": "MarketingTeam"
    //                }
    //            }
    //        }

            string jsonRequestHeader = @"";
            var content = new StringContent(jsonRequestHeader, System.Text.Encoding.UTF8, "application/json");
            try
            {
                // Replace this with your organization's auth url
                string authUri = "http://my-server/api/3.5/auth/signin";
                httpResponse = await httpClient.PostAsync(new Uri(authUri), content);
                string httpResponseMessage = await httpResponse.Content.ReadAsStringAsync();
                JObject responseBody = JObject.Parse(httpResponseMessage);
                authToken = (string)responseBody["credentials"]["token"];
                httpClient.DefaultRequestHeaders.Remove("X-Tableau-Auth");
                httpClient.DefaultRequestHeaders.Add("X-Tableau-Auth", authToken);
                siteId = (string)responseBody["credentials"]["site"]["id"];
            }
            catch (Exception e)
            {
                _logger.LogError("Error while Authentication");
                _logger.LogError(e.Message);
                throw e;
            }
        }
    }
}