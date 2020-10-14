﻿using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TSheetsApi;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Client;
using System.Security;

namespace TSheetIntegration
{
    class Program
    {
        private static string _baseUri = ConfigurationManager.AppSettings.Get("_baseUri");

        private static ConnectionInfo _connection;
        private static IOAuth2 _authProvider;

        private static string _clientId;
        private static string _redirectUri;
        private static string _clientSecret;
        private static string _manualToken;
        static void Main(string[] args)
        {
            // _clientId, _redirectUri, and _clientSecret are needed by the API to connect to your
            // TSheets account.  To get these values for your account, log in to your TSheets account,
            // click on Company Settings -> Add-ons -> API Preferences and use the values for your
            // application. You can specify them through environment variables as shown here, or just
            // paste them into the code here directly.
            Environment.SetEnvironmentVariable("TSHEETS_CLIENTID", ConfigurationManager.AppSettings.Get("TSHEETS_CLIENTID"));
            Environment.SetEnvironmentVariable("TSHEETS_CLIENTSECRET", ConfigurationManager.AppSettings.Get("TSHEETS_CLIENTSECRET"));
            Environment.SetEnvironmentVariable("TSHEETS_REDIRECTURI", ConfigurationManager.AppSettings.Get("TSHEETS_REDIRECTURI"));
            Environment.SetEnvironmentVariable("TSHEETS_MANUALTOKEN", ConfigurationManager.AppSettings.Get("TSHEETS_MANUALTOKEN"));

            _clientId = Environment.GetEnvironmentVariable("TSHEETS_CLIENTID");
            _redirectUri = Environment.GetEnvironmentVariable("TSHEETS_REDIRECTURI");
            _clientSecret = Environment.GetEnvironmentVariable("TSHEETS_CLIENTSECRET");
            _manualToken = Environment.GetEnvironmentVariable("TSHEETS_MANUALTOKEN");

            //NOTE: Set up the ConnectionInfo object which tells the API how to connect to the server
            _connection = new ConnectionInfo(_baseUri, _clientId, _redirectUri, _clientSecret);

            AuthenticateWithManualToken();
            getProjects();
        }

        /// <summary>
        /// Shows how to set up authentication to use a static/manually created access token.
        /// To create a manual auth token, go to the API Add-on preferences in your TSheets account
        /// and click Add Token.
        /// </summary>
        private static void AuthenticateWithManualToken()
        {
            _authProvider = new StaticAuthentication(_manualToken);
        }

        public static void getProjects()
        {
            var url = "https://rest.tsheets.com/api/v1/reports/project";

            var tsheetsApi = new RestClient(_connection, _authProvider);

            var filters = new Dictionary<string, string>();
            filters.Add("start_date", ConfigurationManager.AppSettings.Get("start_date"));
            filters.Add("end_date", ConfigurationManager.AppSettings.Get("end_date"));
            var timesheetData = tsheetsApi.Get(ObjectType.Timesheets, filters);

            var timesheetsObject = JObject.Parse(timesheetData);
            var allTimeSheets = timesheetsObject.SelectTokens("results.timesheets.*");
            var supplemental_data = timesheetsObject.SelectTokens("supplemental_data.jobcodes.*");

            List<AllTimeSheetData> allTimeSheetData = new List<AllTimeSheetData>();
            List<SupplementalData> supplementalData = new List<SupplementalData>();

            //NOTE: Fetch all timesheet data
            foreach (var timesheet in allTimeSheets)
            {
                allTimeSheetData.Add(JsonConvert.DeserializeObject<AllTimeSheetData>(timesheet.ToString()));
            }

            //NOTE: Fetch all supplement data
            foreach (var supplemental in supplemental_data)
            {
                supplementalData.Add(JsonConvert.DeserializeObject<SupplementalData>(supplemental.ToString()));
            }

            string sharepoint_Login = ConfigurationManager.AppSettings.Get("sharepoint_Login");
            string sharepoint_Password = ConfigurationManager.AppSettings.Get("sharepoint_Password");
            var securePassword = new SecureString();
            foreach (char c in sharepoint_Password)
            {
                securePassword.AppendChar(c);
            }

            foreach (var sd in allTimeSheetData)
            {
                if (supplementalData.Where(x => x.id == sd.jobcode_id).Select(x => x.project_id).ToList().Count > 0)
                {
                    long project_id = supplementalData.Where(x => x.id == sd.jobcode_id).Select(x => x.project_id).FirstOrDefault();
                    if (project_id > 0)
                    {
                        string siteUrl = ConfigurationManager.AppSettings.Get("sharepoint_SiteUrl");
                        ClientContext clientContext = new ClientContext(siteUrl);
                        List myList = clientContext.Web.Lists.GetByTitle(ConfigurationManager.AppSettings.Get("sharepoint_ListName"));

                        //NOTE: Check if project id is available in list
                        TimeSpan duration = new TimeSpan();
                        if (!string.IsNullOrWhiteSpace(sd.duration))
                        {
                            duration = TimeSpan.FromSeconds(Convert.ToInt64(sd.duration));

                            string answer = string.Format("{0:D2}h:{1:D2}m:{2:D2}s:{3:D3}ms",
                                            duration.Hours,
                                            duration.Minutes,
                                            duration.Seconds,
                                            duration.Milliseconds);
                        }

                        long ID = CheckItemAlreadyExists(clientContext, sharepoint_Login, securePassword, project_id);
                        if (ID > 0)
                        {
                            ListItem myItem = myList.GetItemById(ID.ToString());
                            myItem["Title"] = sd.id;
                            myItem["user_id"] = sd.user_id;
                            myItem["jobcode_id"] = sd.jobcode_id;
                            myItem["project_id"] = project_id;
                            myItem["Duration"] = duration;

                            myItem.Update();
                            clientContext.ExecuteQuery();
                        }
                        else
                        {
                            ListItemCreationInformation itemInfo = new ListItemCreationInformation();
                            ListItem myItem = myList.AddItem(itemInfo);
                            myItem["Title"] = sd.id;
                            myItem["user_id"] = sd.user_id;
                            myItem["jobcode_id"] = sd.jobcode_id;
                            myItem["project_id"] = project_id;
                            myItem["Duration"] = duration;
                            try
                            {
                                myItem.Update();
                                var onlineCredentials = new SharePointOnlineCredentials(sharepoint_Login, securePassword);
                                clientContext.Credentials = onlineCredentials;
                                clientContext.ExecuteQuery();
                                Console.WriteLine("Item Inserted Successfully project_id: " + project_id);
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine(e.Message);
                            }
                        }

                        //NOTE: Logic for upating PMP sites milestones.
                        GetPMPSitesAndSubSiteTasks(project_id);
                    }
                }
            }
        }

        public static long CheckItemAlreadyExists(ClientContext clientContext, string sharepoint_Login, SecureString securePassword, long project_id)
        {
            long ID = 0;
            List oList = clientContext.Web.Lists.GetByTitle(ConfigurationManager.AppSettings.Get("sharepoint_ListName"));

            CamlQuery camlQuery = new CamlQuery();
            ListItemCollection collListItem = oList.GetItems(camlQuery);

            clientContext.Load(collListItem);

            var onlineCredentials = new SharePointOnlineCredentials(sharepoint_Login, securePassword);
            clientContext.Credentials = onlineCredentials;
            clientContext.ExecuteQuery();

            foreach (ListItem oListItem in collListItem)
            {
                if (project_id == Convert.ToInt64(oListItem["project_id"]))
                {
                    return oListItem.Id;
                }
                //Console.WriteLine("ID: {0} \nTitle: {1} \nBody: {2}", oListItem.Id, oListItem["project_id"], oListItem["Body"]);
            }
            return ID;
        }

        public static void GetPMPSitesAndSubSiteTasks(long project_id)
        {
            string siteUrl = "https://leonlebeniste.sharepoint.com/sites/PMP";
            ClientContext clientContext = new ClientContext(siteUrl);

            long ID = 0;
            List oList = clientContext.Web.Lists.GetByTitle("LL Projects List");

            CamlQuery camlQuery = new CamlQuery();
            ListItemCollection collListItem = oList.GetItems(camlQuery);

            clientContext.Load(collListItem);

            string sharepoint_Login = ConfigurationManager.AppSettings.Get("sharepoint_Login_PMP");
            string sharepoint_Password = ConfigurationManager.AppSettings.Get("sharepoint_Password_PMP");
            var securePassword = new SecureString();
            foreach (char c in sharepoint_Password)
            {
                securePassword.AppendChar(c);
            }

            var onlineCredentials = new SharePointOnlineCredentials(sharepoint_Login, securePassword);
            clientContext.Credentials = onlineCredentials;
            clientContext.ExecuteQuery();

            foreach (ListItem oListItem in collListItem)
            {
                if (project_id == Convert.ToInt64(oListItem["ProjID"]))
                {
                    string subSiteURL = ((Microsoft.SharePoint.Client.FieldUrlValue)oListItem["SiteURL"]).Url;

                    //NOTE: Get Sub Site Tasks items.
                    GetPMPSubSiteTaskLists(subSiteURL, sharepoint_Login, securePassword);
                }
            }
        }

        public static void GetPMPSubSiteTaskLists(string siteUrl, string sharepoint_Login, SecureString securePassword)
        {
            ClientContext clientContext = new ClientContext(siteUrl);

            long ID = 0;
            List oList = clientContext.Web.Lists.GetByTitle("Schedule");

            CamlQuery camlQuery = new CamlQuery();
            ListItemCollection collListItem = oList.GetItems(camlQuery);

            clientContext.Load(collListItem);

            var onlineCredentials = new SharePointOnlineCredentials(sharepoint_Login, securePassword);
            clientContext.Credentials = onlineCredentials;
            clientContext.ExecuteQuery();

            foreach (ListItem oListItem in collListItem)
            {
                //if (project_id == Convert.ToInt64(oListItem["ProjID"]))
                //{
                //    string siteURL = ((Microsoft.SharePoint.Client.FieldUrlValue)oListItem["SiteURL"]).Url;
                //    GetPMPSubSiteTaskLists(siteURL);
                //}
                //Console.WriteLine("ID: {0} \nTitle: {1} \nBody: {2}", oListItem.Id, oListItem["project_id"], oListItem["Body"]);
            }
        }

    }
}
