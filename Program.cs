using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Diagnostics;

//1.Build Console App
//2.Publish to Azure
//3.On AAD, register App, redirect Uri address
//4.On AAD, Set 'Client Secret'
//5-1.On AAD, Set 'API permissions'> Microsoft Graph > User.Read, User.Read.All > Grant Admin Consent
//5-2.On AAD, Set 'API permissions'> Microsoft Graph > Group.Read, Group.Read.All > Grant Admin Consent
//5-3.On AAD, Set 'API permissions'> Microsoft Graph > GroupMember.Read, GroupMember.Read.All > Grant Admin Consent


namespace ConsoleGraphTest
{
    class Program
    {
        private static GraphServiceClient _graphServiceClient;
        private static HttpClient _httpClient;

        static void Main(string[] args)
        {
            // Load appsettings.json - IConfigurationRoot
            var config = LoadAppSettings();
            if (null == config)
            {
                Console.WriteLine("Missing or invalid appsettings.json file. Please see README.md for configuration instructions.");
                return;
            }

            //WAY1: Query using Graph SDK (preferred when possible)
            GraphServiceClient graphClient = GetAuthenticatedGraphClient(config);

            List<QueryOption> options = new List<QueryOption>
            {
                new QueryOption("$top", "1")
            };

            //1.User
            var graphResult1 = graphClient.Users.Request(options).GetAsync().Result;
            Console.WriteLine("Graph SDK Result");
            Console.WriteLine(graphResult1[0].DisplayName);
            Debug.WriteLine("Graph SDK Result");
            Debug.WriteLine(graphResult1[0].DisplayName);


            //2.Group
            var allAdGroups = new List<AdGroupConfig>();
            var groups = graphClient.Groups.Request().GetAsync().Result;
            if (groups.Count > 0)
            {
                allAdGroups.AddRange(groups.Select(a => new AdGroupConfig() { GroupId = a.Id, GroupName = a.DisplayName }));
                allAdGroups.ForEach(a => Debug.WriteLine($"groupId:{a.GroupId}, GroupName:{a.GroupName}"));
            }


            //WAY2: Direct query using HTTPClient (for beta endpoint calls or not available in Graph SDK)
            HttpClient httpClient = GetAuthenticatedHTTPClient(config);
            Uri Uri = new Uri("https://graph.microsoft.com/v1.0/users?$top=1");
            var httpResult = httpClient.GetStringAsync(Uri).Result;

            Console.WriteLine("HTTP Result");
            Console.WriteLine(httpResult);
            Debug.WriteLine("HTTP Result");
            Debug.WriteLine(httpResult);
        }

        private static GraphServiceClient GetAuthenticatedGraphClient(IConfigurationRoot config)
        {
            var authenticationProvider = CreateAuthorizationProvider(config);
            _graphServiceClient = new GraphServiceClient(authenticationProvider);
            return _graphServiceClient;
        }

        private static HttpClient GetAuthenticatedHTTPClient(IConfigurationRoot config)
        {
            var authenticationProvider = CreateAuthorizationProvider(config);
            _httpClient = new HttpClient(new AuthHandler(authenticationProvider, new HttpClientHandler()));
            return _httpClient;
        }

        private static IAuthenticationProvider CreateAuthorizationProvider(IConfigurationRoot config)
        {
            var clientId = config["AzureAd:ClientId"];
            var clientSecret = config["AzureAd:ClientSecret"];
            var redirectUri = config["AzureAd:RedirectUri"];
            var authority = $"https://login.microsoftonline.com/{config["AzureAd:TenantId"]}/v2.0";

            //this specific scope means that application will default to what is defined in the application registration rather than using dynamic scopes
            List<string> scopes = new List<string>();
            scopes.Add("https://graph.microsoft.com/.default");

            var cca = ConfidentialClientApplicationBuilder.Create(clientId)
                                                    .WithAuthority(authority)
                                                    .WithRedirectUri(redirectUri)
                                                    .WithClientSecret(clientSecret)
                                                    .Build();
            return new MsalAuthenticationProvider(cca, scopes.ToArray());
        }

        private static IConfigurationRoot LoadAppSettings()
        {
            try
            {
                var builder = new ConfigurationBuilder()
                .SetBasePath(System.IO.Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);

                IConfigurationRoot config = builder.Build();

                Debug.WriteLine(config["AzureAd:ClientId"]);
                Debug.WriteLine(config["AzureAd:ClientSecret"]);
                Debug.WriteLine(config["AzureAd:TenantId"]);
                Debug.WriteLine(config["AzureAd:RedirectUri"]);
                Debug.WriteLine(config["AzureAd:Domain"]);

                // Validate required settings
                if (string.IsNullOrEmpty(config["AzureAd:ClientId"]) ||
                    string.IsNullOrEmpty(config["AzureAd:ClientSecret"]) ||
                    string.IsNullOrEmpty(config["AzureAd:TenantId"]) ||
                    string.IsNullOrEmpty(config["AzureAd:RedirectUri"]) ||
                    string.IsNullOrEmpty(config["AzureAd:Domain"]))
                {
                    return null;
                }

                return config;
            }
            catch (System.IO.FileNotFoundException)
            {
                return null;
            }
        }

        public class AdGroupConfig
        {
            public string GroupId { get; set; }
            public string GroupName { get; set; }
        }



    }
}