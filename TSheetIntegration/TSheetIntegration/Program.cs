using Newtonsoft.Json;
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
        private static int count = 0;
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
            //getProjects();
            GetAllProjectIdForProjectId();
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

        public static void GetAllProjectIdForProjectId()
        {
            int currentPage = 1;
            bool moreData = true;
            List<string> projectNames = new List<string>();
            List<Projects> projects = new List<Projects>();
            var tsheetsApi = new RestClient(_connection, _authProvider);
            while (moreData)
            {
                var filters = new Dictionary<string, string>();
                filters.Add("parent_ids", "0");
                filters["per_page"] = "50";
                filters["page"] = currentPage.ToString();

                var projectData = tsheetsApi.Get(ObjectType.Jobcodes, filters);
                var projectDataObj = JObject.Parse(projectData);
                var ienumProjectData = projectDataObj.SelectTokens("results.jobcodes.*");
                foreach (var ie in ienumProjectData)
                {
                    projects.Add(JsonConvert.DeserializeObject<Projects>(ie.ToString()));
                }
                // see if we have more pages to retrieve
                moreData = bool.Parse(projectDataObj.SelectToken("more").ToString());

                // increment to the next page
                currentPage++;
            }

            //NOTE: Get all the projects lists
            string PMPSiteUrl = "https://leonlebeniste.sharepoint.com/sites/PMP";
            ClientContext clientContext = new ClientContext(PMPSiteUrl);

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

            foreach (var project in projects)
            {
                //NOTE: Process updating project id

                string projectName = project.name.Substring(0, 5);
                if (!projectNames.Contains(projectName))
                {
                    projectNames.Add(projectName);

                    foreach (ListItem oListItem in collListItem)
                    {
                        if (projectName == Convert.ToString(oListItem["ProjectNumber"]))
                        {
                            if (project.id != Convert.ToInt64(oListItem["ProjID"]))
                            {
                                //NOTE: Update Project ID in list.
                                ListItem myItem = oList.GetItemById(Convert.ToString(oListItem["ID"]));
                                myItem["ProjID"] = project.id;
                                try
                                {
                                    myItem.Update();
                                    clientContext.Credentials = onlineCredentials;
                                    clientContext.ExecuteQuery();
                                    Console.WriteLine("Project ID Successfully Update for: " + projectName);
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine(e.Message);
                                }
                            }
                            //NOTE: Process updating timesheet hours.
                            GetAllJobCodeIdForProjectId(project.id);
                        }
                    }
                }
            }
        }

        public static void GetAllJobCodeIdForProjectId(long projectId)
        {
            //long projectId = 56135257;
            var tsheetsApi = new RestClient(_connection, _authProvider);
            var filters = new Dictionary<string, string>();
            filters.Add("parent_ids", projectId.ToString());

            var milestoneData = tsheetsApi.Get(ObjectType.Jobcodes, filters);
            var milestoneDataObj = JObject.Parse(milestoneData);
            var ienumJCData = milestoneDataObj.SelectTokens("results.jobcodes.*");
            List<MilestoneData> allMilestoneData = new List<MilestoneData>();
            foreach (var ie in ienumJCData)
            {
                allMilestoneData.Add(JsonConvert.DeserializeObject<MilestoneData>(ie.ToString()));
            }

            foreach (var msData in allMilestoneData)
            {
                GetAllTimeSheetDataForJobCodeId(projectId, msData.id, msData.name, tsheetsApi);
            }
        }

        public static void GetAllTimeSheetDataForJobCodeId(long projectId, long jobCodeId, string taskName, RestClient tsheetsApi)
        {
            int currentPage = 1;
            bool moreData = true;
            List<AllTimeSheetData> allTimeSheetData = new List<AllTimeSheetData>();
            while (moreData)
            {
                var filters = new Dictionary<string, string>();
                filters.Add("start_date", ConfigurationManager.AppSettings.Get("start_date"));
                //filters.Add("end_date", ConfigurationManager.AppSettings.Get("end_date"));
                filters.Add("end_date", DateTime.Now.ToString("yyyy-MM-dd"));
                filters.Add("jobcode_ids", jobCodeId.ToString());
                filters["per_page"] = "50";
                filters["page"] = currentPage.ToString();
                var timesheetData = tsheetsApi.Get(ObjectType.Timesheets, filters);
                var timesheetsObject = JObject.Parse(timesheetData);
                var allTimeSheets = timesheetsObject.SelectTokens("results.timesheets.*");
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

                // see if we have more pages to retrieve
                moreData = bool.Parse(timesheetsObject.SelectToken("more").ToString());

                // increment to the next page
                currentPage++;
            }
            GetPMPSitesAndSubSiteURL(projectId, taskName, allTimeSheetData);
            count = 0;
        }

        public static void GetPMPSitesAndSubSiteURL(long project_id, string taskName, List<AllTimeSheetData> allMilestoneItems)
        {
            string siteUrl = "https://leonlebeniste.sharepoint.com/sites/PMP";
            ClientContext clientContext = new ClientContext(siteUrl);

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
                    GetAndSetDurationOnPMPSubSiteTaskLists(subSiteURL, sharepoint_Login, securePassword, taskName, allMilestoneItems);
                }
            }
        }

        public static void GetAndSetDurationOnPMPSubSiteTaskLists(string siteUrl, string sharepoint_Login, SecureString securePassword, string taskName, List<AllTimeSheetData> allMilestoneItems)
        {
            ClientContext clientContext = new ClientContext(siteUrl);

            List oList = clientContext.Web.Lists.GetByTitle("Schedule");

            CamlQuery camlQuery = new CamlQuery();
            ListItemCollection collListItem = oList.GetItems(camlQuery);

            clientContext.Load(collListItem);

            var onlineCredentials = new SharePointOnlineCredentials(sharepoint_Login, securePassword);
            clientContext.Credentials = onlineCredentials;
            clientContext.ExecuteQuery();

            //var groupItem = allTimeSheetData.Where(x => x.jobcode_id == spItem.Select(y => y.id).FirstOrDefault()).GroupBy(x => x.id);

            float installation = 0; float projectManagement = 0; float fabrication = 0; float preProduction = 0;
            float installationVal = 0; float projectManagementVal = 0;
            float fabricationVal = 0; float preProductionVal = 0;

            foreach (var item in allMilestoneItems)
            {
                if (item.customfields.SecondColumn == "Installation")
                {
                    installation = installation + Convert.ToInt64(item.duration);
                }
                else if (item.customfields.SecondColumn == "Project Management")
                {
                    projectManagement = projectManagement + Convert.ToInt64(item.duration);
                }
                else if (item.customfields.SecondColumn == "Fabrication")
                {
                    fabrication = fabrication + Convert.ToInt64(item.duration);
                }
                else if (item.customfields.SecondColumn == "Pre Production")
                {
                    preProduction = preProduction + Convert.ToInt64(item.duration);
                }
            }

            float installhours = (float)System.Math.Round(installation / 3600, 2);
            installationVal = installhours;

            float projectMhours = (float)System.Math.Round(projectManagement / 3600, 2);
            projectManagementVal = projectMhours;

            float fabhours = (float)System.Math.Round(fabrication / 3600, 2);
            fabricationVal = fabhours;

            float prePhours = (float)System.Math.Round(preProduction / 3600, 2);
            preProductionVal = prePhours;

            foreach (ListItem oListItem in collListItem)
            {
                if (Convert.ToString(oListItem["Title"]) == taskName)
                {
                    ListItemCreationInformation itemInfo = new ListItemCreationInformation();
                    ListItem myItem = oList.GetItemById(Convert.ToString(oListItem["ID"]));
                    myItem["Actual_x0020_Install"] = installationVal;
                    myItem["Actual_x0020_Project_x0020_Manag"] = projectManagementVal;
                    myItem["Actual_x0020_Fabrication"] = fabricationVal;
                    myItem["Actual_x0020_Pre_x0020_Productio"] = preProductionVal;
                    try
                    {
                        //myItem.Update();
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
