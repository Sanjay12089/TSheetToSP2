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
using System.Drawing.Drawing2D;
using Microsoft.ProjectServer.Client;

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
            //filters.Add("end_date", ConfigurationManager.AppSettings.Get("end_date"));
            filters.Add("end_date", DateTime.Now.ToString("yyyy-MM-dd"));
            var timesheetData = tsheetsApi.Get(ObjectType.Timesheets, filters);
            var timesheetsObject = JObject.Parse(timesheetData);
            var allTimeSheets = timesheetsObject.SelectTokens("results.timesheets.*");
            var supplemental_data = timesheetsObject.SelectTokens("supplemental_data.jobcodes.*");

            List<AllTimeSheetData> allTimeSheetData = new List<AllTimeSheetData>();
            List<SupplementalData> supplementalData = new List<SupplementalData>();

            //NOTE: Fetch all timesheet data
            int count = 0;
            foreach (var timesheet in allTimeSheets)
            {
                allTimeSheetData.Add(JsonConvert.DeserializeObject<AllTimeSheetData>(timesheet.ToString()));
                int cs = 0;
                foreach (var item in timesheet["customfields"])
                {
                    if (cs == 0)
                        allTimeSheetData[count].customfields.FirstColumn = item.First.ToString();
                    if (cs == 1)
                        allTimeSheetData[count].customfields.SecondColumn = item.First.ToString();
                    if (cs == 2)
                        allTimeSheetData[count].customfields.ThirdColumn = item.First.ToString();
                    if (cs == 3)
                        allTimeSheetData[count].customfields.FourthColumn = item.First.ToString();
                    if (cs == 4)
                        allTimeSheetData[count].customfields.FifthColumn = item.First.ToString();
                    cs++;
                }
                count++;
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

            foreach (var td in allTimeSheetData)
            {
                List<SupplementalData> spItem = supplementalData.Where(x => x.id == td.jobcode_id).ToList();
                if (spItem.Count > 0)
                {
                    long project_id = supplementalData.Where(x => x.id == td.jobcode_id).Select(x => x.project_id).FirstOrDefault();
                    if (project_id > 0)
                    {
                        #region trial tenant list

                        //string siteUrl = ConfigurationManager.AppSettings.Get("sharepoint_SiteUrl");
                        //ClientContext clientContext = new ClientContext(siteUrl);
                        //List myList = clientContext.Web.Lists.GetByTitle(ConfigurationManager.AppSettings.Get("sharepoint_ListName"));

                        //NOTE: Check if project id is available in list
                        //TimeSpan duration = new TimeSpan();
                        //if (!string.IsNullOrWhiteSpace(sd.duration))
                        //{
                        //    duration = TimeSpan.FromSeconds(Convert.ToInt64(sd.duration));

                        //    string answer = string.Format("{0:D2}h:{1:D2}m:{2:D2}s:{3:D3}ms",
                        //                    duration.Hours,
                        //                    duration.Minutes,
                        //                    duration.Seconds,
                        //                    duration.Milliseconds);
                        //}

                        //long ID = CheckItemAlreadyExists(clientContext, sharepoint_Login, securePassword, project_id);
                        //if (ID > 0)
                        //{
                        //    ListItem myItem = myList.GetItemById(ID.ToString());
                        //    myItem["Title"] = sd.id;
                        //    myItem["user_id"] = sd.user_id;
                        //    myItem["jobcode_id"] = sd.jobcode_id;
                        //    myItem["project_id"] = project_id;
                        //    myItem["Duration"] = duration;

                        //    myItem.Update();
                        //    clientContext.ExecuteQuery();
                        //}
                        //else
                        //{
                        //    ListItemCreationInformation itemInfo = new ListItemCreationInformation();
                        //    ListItem myItem = myList.AddItem(itemInfo);
                        //    myItem["Title"] = sd.id;
                        //    myItem["user_id"] = sd.user_id;
                        //    myItem["jobcode_id"] = sd.jobcode_id;
                        //    myItem["project_id"] = project_id;
                        //    myItem["Duration"] = duration;
                        //    try
                        //    {
                        //        myItem.Update();
                        //        var onlineCredentials = new SharePointOnlineCredentials(sharepoint_Login, securePassword);
                        //        clientContext.Credentials = onlineCredentials;
                        //        clientContext.ExecuteQuery();
                        //        Console.WriteLine("Item Inserted Successfully project_id: " + project_id);
                        //    }
                        //    catch (Exception e)
                        //    {
                        //        Console.WriteLine(e.Message);
                        //    }
                        //}

                        #endregion

                        string taskName = spItem.Select(x => x.name).FirstOrDefault();

                        //NOTE: Logic for upating PMP sites milestones.
                        GetPMPSitesAndSubSiteTasks(project_id, taskName, allTimeSheetData, spItem);
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

        public static void GetPMPSitesAndSubSiteTasks(long project_id, string taskName, List<AllTimeSheetData> allTimeSheetData, List<SupplementalData> spItem)
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
                    GetPMPSubSiteTaskLists(subSiteURL, sharepoint_Login, securePassword, taskName, allTimeSheetData, spItem);
                }
            }
        }

        public static void GetPMPSubSiteTaskLists(string siteUrl, string sharepoint_Login, SecureString securePassword, string taskName, List<AllTimeSheetData> allTimeSheetData, List<SupplementalData> spItem)
        {
            ClientContext clientContext = new ClientContext(siteUrl);

            List oList = clientContext.Web.Lists.GetByTitle("Schedule");

            CamlQuery camlQuery = new CamlQuery();
            ListItemCollection collListItem = oList.GetItems(camlQuery);

            clientContext.Load(collListItem);

            var onlineCredentials = new SharePointOnlineCredentials(sharepoint_Login, securePassword);
            clientContext.Credentials = onlineCredentials;
            clientContext.ExecuteQuery();

            var groupItem = allTimeSheetData.Where(x => x.jobcode_id == spItem.Select(y => y.id).FirstOrDefault()).GroupBy(x => x.id);

            long installation = 0; long projectManagement = 0; long fabrication = 0; long preProduction = 0;
            string installationVal = string.Empty; string projectManagementVal = string.Empty;
            string fabricationVal = string.Empty; string preProductionVal = string.Empty;

            foreach (var item in groupItem)
            {
                foreach (var it in item)
                {
                    if (it.customfields.SecondColumn == "Installation")
                    {
                        installation = installation + Convert.ToInt64(it.duration);
                    }
                    else if (it.customfields.SecondColumn == "Project Management")
                    {
                        projectManagement = projectManagement + Convert.ToInt64(it.duration);
                    }
                    else if (it.customfields.SecondColumn == "Fabrication")
                    {
                        fabrication = fabrication + Convert.ToInt64(it.duration);
                    }
                    else if (it.customfields.SecondColumn == "Pre Production")
                    {
                        preProduction = preProduction + Convert.ToInt64(it.duration);
                    }
                }
            }

            TimeSpan duration = new TimeSpan();
            duration = TimeSpan.FromSeconds(Convert.ToInt64(installation));
            installationVal = string.Format("{0:D2}h:{1:D2}m:{2:D2}s:{3:D3}ms",
                            duration.Hours,
                            duration.Minutes,
                            duration.Seconds,
                            duration.Milliseconds);

            duration = new TimeSpan();
            duration = TimeSpan.FromSeconds(Convert.ToInt64(projectManagement));
            projectManagementVal = string.Format("{0:D2}h:{1:D2}m:{2:D2}s:{3:D3}ms",
                            duration.Hours,
                            duration.Minutes,
                            duration.Seconds,
                            duration.Milliseconds);

            duration = new TimeSpan();
            duration = TimeSpan.FromSeconds(Convert.ToInt64(fabrication));
            fabricationVal = string.Format("{0:D2}h:{1:D2}m:{2:D2}s:{3:D3}ms",
                            duration.Hours,
                            duration.Minutes,
                            duration.Seconds,
                            duration.Milliseconds);

            duration = new TimeSpan();
            duration = TimeSpan.FromSeconds(Convert.ToInt64(preProduction));
            preProductionVal = string.Format("{0:D2}h:{1:D2}m:{2:D2}s:{3:D3}ms",
                            duration.Hours,
                            duration.Minutes,
                            duration.Seconds,
                            duration.Milliseconds);

            foreach (ListItem oListItem in collListItem)
            {
                if (oListItem["Title"] == taskName)
                {
                    ListItemCreationInformation itemInfo = new ListItemCreationInformation();
                    ListItem myItem = oList.AddItem(itemInfo);
                    myItem["3DActual_x0020_Install"] = installationVal;
                    myItem["3DActual_x0020_Project_x0020_Manag"] = projectManagementVal;
                    myItem["3DActual_x0020_Fabrication"] = fabricationVal;
                    myItem["3DActual_x0020_Pre_x0020_Productio"] = preProductionVal;
                    try
                    {
                        myItem.Update();
                        clientContext.Credentials = onlineCredentials;
                        //clientContext.ExecuteQuery();
                        Console.WriteLine("Item Updated Successfully name: " + taskName);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                    }
                }
            }
        }

    }
}
