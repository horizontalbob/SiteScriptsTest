using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Newtonsoft.Json;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Text;

namespace SiteScriptTestFramework
{
    internal class Program
    {
        private static readonly string clientId = "YOUR-CLIENTID";
        private static readonly string clientSecret = "YOUR-CLIENT-SECRET";
        private static readonly string tenantId = "YOUR-TENANTID";
        private static string siteUrl = "YOUR-SITE-URL";
        private static string basesiteUrl = "YOUR-BASE-SITE-URL";
        private static string username = "YOUR-USER-EMAIL";
        private static string resource = "https://graph.microsoft.com";
        private static string certThumbprint = "YOUR-CERT-THUMBPRINT";
        private static string[] scopes = new string[] { $"{basesiteUrl}/.default" };
        private static string authority = $"https://login.microsoftonline.com/{tenantId}";
        private static string accessToken = "ACCESS_TOKEN";
        private static IConfidentialClientApplication app;
        private static AuthenticationResult result;
        private static string siteScriptJson;

        #region EndPoints

        private static string getScriptsUrl = "/_api/Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScripts";
        private static string deleteScriptUrl = "/_api/Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility.DeleteSiteScript";
        private static string createSiteUrl = "/_api/SPSiteManager/create";
        private static string createDesignTemplateUrl = "/_api/$metadata#Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteDesignMetadata";
        private static string getDesignTemplatesUrl = "/_api/Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns";

        #endregion

        public static async Task Main(string[] args)
        {
            await DoSomeAuth();

            string selection = null;
            while (selection != "0")
            {
                Console.WriteLine("Select a menu option or 0 to exit:");
                Console.WriteLine("1) List all current site scripts");
                Console.WriteLine("2) Delete all current site scripts");
                Console.WriteLine("3) Bulk create site scripts");
                Console.WriteLine("4) Create site from site script");
                selection = Console.ReadLine();
                switch (selection)
                {
                    case "1":
                        await ListAllScripts();
                        break;
                    case "2":
                        await DeleteAllScripts();
                        break;
                    case "3":
                        Console.WriteLine("Amount of scripts to create: ");
                        var amount = Console.ReadLine();
                        await LoadMultipleScripts(siteScriptJson, Int32.Parse(amount));
                        break;
                    case "4":
                        await CreateSite();
                        break;
                    default:
                        selection = null;
                        Console.WriteLine("Unknown input");
                        break;
                    case "0":
                        Console.WriteLine("Bye!");
                        break;
                }
            }


            await CreateSite();

            Console.WriteLine("Press any key to close");
            Console.ReadKey();
        }

        private static async Task DoSomeAuth()
        {
            X509Certificate2 certificate = GetCertificateFromStore(certThumbprint);

            if (certificate == null)
            {
                Console.WriteLine("Certificate not found.");
                return;
            }

            app = ConfidentialClientApplicationBuilder.Create(clientId)
                .WithCertificate(certificate)
                .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}"))
                .Build();

            result = await app.AcquireTokenForClient(scopes).ExecuteAsync();

            accessToken = result.AccessToken;
        }

        private static void SetupSiteJson()
        {
            var siteScript = new SiteScript();

            var siteScriptActions = new List<Action>
            {
                new Action
                {
                    verb = "createSPList",
                    listName = "Customer Tracking",
                    templateType = 100,
                    subactions = new List<Subaction>
                    {
                        new Subaction
                        {
                            verb = "setDescription",
                            description = "List of Customers and Orders"
                        }
                    }
                },
            };

            siteScript.actions = siteScriptActions;
            siteScript.schema = "https://developer.microsoft.com/json-schemas/sp/site-design-script-actions.schema.json";
            siteScript.bindData = new BindData();

            siteScriptJson = JsonConvert.SerializeObject(siteScript);
        }

        private static async Task<GetDesignTemplatesResponse> ListSiteDesignTemplates()
        {
            var scripts = await SendRequest(null, getScriptsUrl, "List site design templates");
            Console.WriteLine(scripts);
            return JsonConvert.DeserializeObject<GetDesignTemplatesResponse>(scripts);
        }

        private static async Task CreateDesignTemplate()
        {
            var scriptResult = await ListAllScripts();

            if (scriptResult.value.Length < 1)
            {
                await LoadMultipleScripts(siteScriptJson);
            }

            scriptResult = await ListAllScripts();

            var siteDesignObject = new CreateSiteDesignObject
            {
                Description = "Test site template",
                SiteScriptIds = new[] { scriptResult.value.FirstOrDefault().Id },
                Title = "My lovely site",
                WebTemplate = "64"
            };

            var siteDesignJson = JsonConvert.SerializeObject(siteDesignObject);

            await SendRequest(siteDesignJson, createDesignTemplateUrl, "Create site design template");
        }

        private static async Task CreateSite()
        {
            //await CreateDesignTemplate();
            var scriptResult = await ListSiteDesignTemplates();

            //"f6cc5403-0d63-442e-96c0-285923709ffc",
            var createSiteRequest = new CreateSiteRequest
            {
                request = new CreateSiteBody
                {
                    Title = "Site script created site",
                    Url = $"{basesiteUrl}/sites/TestSite{new Random().Next(10, 99)}",
                    Lcid = 1033,
                    ShareByEmailEnabled = false,
                    //Classification = "Low Business Impact",
                    Description = "Description",
                    WebTemplate = "SITEPAGEPUBLISHING#0",
                    SiteDesignId = scriptResult.value.FirstOrDefault().Id,
                    Owner = username,
                    WebTemplateExtensionId = "00000000-0000-0000-0000-000000000000"
                }
            };

            Console.WriteLine($"CreateSiteRequest: {createSiteRequest.request.Url}");

            await SendRequest(JsonConvert.SerializeObject(createSiteRequest), createSiteUrl, $"Creating site: {createSiteRequest.request.Url}");
        }

        private static async Task DeleteAllScripts()
        {
            var scripts = await SendRequest(null, getScriptsUrl, "List site scripts");
            var scriptObjects = JsonConvert.DeserializeObject<GetScriptsResponse>(scripts);
            var siteScriptids = scriptObjects.value.Select(x => new DeleteId { id = x.Id });

            foreach (var id in siteScriptids)
            {
                await SendRequest(JsonConvert.SerializeObject(id), deleteScriptUrl, $"Deleted site script {id.id}");
            }
        }

        private static async Task LoadMultipleScripts(string siteScriptJson, int count = 1)
        {
            for (var i = 0; i < count; i++)
            {
                var createSiteScriptUrl = $"/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteScript(Title=@title)?@title='New script{i}'";

                await SendRequest(siteScriptJson, createSiteScriptUrl, "Creating site");
                Console.WriteLine(i);
            }
        }

        private static async Task<GetScriptsResponse> ListAllScripts()
        {
            var scripts = await SendRequest(null, getScriptsUrl, "List site scripts");
            Console.WriteLine(scripts);
            return JsonConvert.DeserializeObject<GetScriptsResponse>(scripts);
        }

        private static X509Certificate2 GetCertificateFromStore(string thumbprint)
        {
            X509Certificate2 certificate = null;
            X509Store store = new X509Store(StoreLocation.CurrentUser);

            try
            {
                store.Open(OpenFlags.ReadOnly);

                X509Certificate2Collection certCollection = store.Certificates.Find(
                    X509FindType.FindByThumbprint, thumbprint, false);

                if (certCollection.Count > 0)
                {
                    certificate = certCollection[0];
                }
            }
            finally
            {
                store.Close();
            }

            return certificate;
        }

        private static async Task<string> SendRequest(string siteScriptJson, string url, string action)
        {
            HttpResponseMessage response = null;
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, $"{basesiteUrl}{url}");

                    if (siteScriptJson != null)
                    {
                        request.Content = new StringContent(siteScriptJson, Encoding.UTF8, "application/json");
                    }
                    try
                    {
                        response = await client.SendAsync(request);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                        return null;
                    }

                    if (response.IsSuccessStatusCode)
                    {
                        string resultResponse = await response.Content.ReadAsStringAsync();
                        Console.WriteLine(action);
                        return resultResponse;
                    }
                    else
                    {
                        string error = await response.Content.ReadAsStringAsync();
                        Console.WriteLine("Error creating Site Script: " + error);
                        Console.WriteLine("Status Code: " + response.StatusCode);
                        Console.WriteLine("Reason Phrase: " + response.ReasonPhrase);
                        return error;
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                return e.ToString();
            }
        }
    }

    #region Object models

    public class GetScriptsResponse
    {
        public string odatacontext { get; set; }
        public Value[] value { get; set; }
    }

    public class Value
    {
        public object Content { get; set; }
        public object Description { get; set; }
        public string Id { get; set; }
        public string Title { get; set; }
        public int Version { get; set; }
    }

    public class DeleteId
    {
        public string id { get; set; }
    }


    public class CreateSiteRequest
    {
        public CreateSiteBody request { get; set; }
    }

    public class CreateSiteBody
    {
        public string Title { get; set; }
        public string Url { get; set; }
        public int Lcid { get; set; }
        public bool ShareByEmailEnabled { get; set; }
        public string Classification { get; set; }
        public string Description { get; set; }
        public string WebTemplate { get; set; }
        public string SiteDesignId { get; set; }
        public string Owner { get; set; }
        public string WebTemplateExtensionId { get; set; }
    }

    public class CreateList
    {
        public __Metadata __metadata { get; set; }
        public bool AllowContentTypes { get; set; }
        public int BaseTemplate { get; set; }
        public bool ContentTypesEnabled { get; set; }
        public string Description { get; set; }
        public string Title { get; set; }
    }

    public class __Metadata
    {
        public string type = "SP.List";
    }


    public class CreateSiteDesignObject
    {
        public string odatacontext { get; set; }
        public string Description { get; set; }
        public string PreviewImageAltText { get; set; }
        public string PreviewImageUrl { get; set; }
        public string[] SiteScriptIds { get; set; }
        public string Title { get; set; }
        public string WebTemplate { get; set; }
        public string Id { get; set; }
        public int Version { get; set; }
    }


    public class GetDesignTemplatesResponse
    {
        public string odatacontext { get; set; }
        public GetDesignTemplatesValue[] value { get; set; }
    }

    public class GetDesignTemplatesValue
    {
        public string Description { get; set; }
        public bool IsDefault { get; set; }
        public string PreviewImageAltText { get; set; }
        public string PreviewImageUrl { get; set; }
        public string[] SiteScriptIds { get; set; }
        public string Title { get; set; }
        public string WebTemplate { get; set; }
        public string Id { get; set; }
        public int Version { get; set; }
    }



    #endregion
}
