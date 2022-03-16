using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace MSALSharePointPermissions
{
    class Program
    { 
        static IPublicClientApplication publicClientApp;

        static async Task Main(string[] args)
        {
            // This is the app reg id I have registered as a multi-tenant app. Feel free to use it or stand up your own app reg
            // by following the setup instructions in the readme.md
            const string clientId = "8fd1129c-f58e-4d41-b4b3-177f3df1a106";
            const string sharepointTenantUrl = "https://YOUR-TENANT.sharepoint.com"; // <- REPLACE YOUR SHAREPOINT TENANT HERE

            Console.WriteLine("MSAL SharePoint permissions test app");
            Console.WriteLine("You should be prompted to login into M365, if you have no existing consent for this\napp you should be presented initially just for SharePoint scopes.\n\n");

            publicClientApp = PublicClientApplicationBuilder.Create(clientId)
                .WithRedirectUri("http://localhost")
                .Build();

            var accounts = await publicClientApp.GetAccountsAsync();

            // First let's make a call to SharePoint REST API to get the user profile just specifying we need the SharePoint allsites.manage scope (we expect this to fail)
            string[] sharepointScopes = new string[] { $"{sharepointTenantUrl}/allsites.manage" };
            await MakeSharePointRestApiCallToGetUserProfile(sharepointTenantUrl, accounts, sharepointScopes);

            // Now let's get the user to consent to the Graph API user.read permission
            Console.WriteLine($"Getting user to consent to the Graph API 'user.read' scope");
            string[] graphScopes = new string[] { $"https://graph.microsoft.com/user.read" };
            AuthenticationResult graphAuthResult = await AuthenticateWithAzureADAsync(graphScopes, accounts.FirstOrDefault());

            // Clear the MSAL cache
            Console.WriteLine($"Clearing the MSAL cache to force MSAL to go back to Azure to get the SharePoint refesh and access token again now we've consented to the Graph permission");
            while (accounts.Any())
            {
                await publicClientApp.RemoveAsync(accounts.First());
                accounts = (await publicClientApp.GetAccountsAsync()).ToList();
            }

            // Need to give identity service a while after granting consent on the Graph permission or the next step doesn't work
            Console.WriteLine("Waiting 30 secs to give Azure identity service time to acknowledge the new Graph permission granted");
            await Task.Delay(30000);

            // Make exactly the same call to the SharePoint REST API to get the user profile (that failed earlier). The only difference is that in between the user has granted the Graph user.read scope.
            await MakeSharePointRestApiCallToGetUserProfile(sharepointTenantUrl, accounts, sharepointScopes);

            Console.WriteLine("If you got here and saw the user profile data come back we've proven access the the user profile (in the SharePoint REST API) is driven by the Graph user.read scope!");
            Console.WriteLine("\nPress any key to finish");
            Console.ReadKey();
        }

        private static async Task MakeSharePointRestApiCallToGetUserProfile(string sharepointTenantUrl, IEnumerable<IAccount> accounts, string[] sharepointScopes)
        {
            Console.WriteLine($"Getting an access token for the SharePoint resource requesting 'allsites.manage' scope");
            AuthenticationResult sharePointAuthResult = await AuthenticateWithAzureADAsync(sharepointScopes, accounts.FirstOrDefault());

            // Now make call to SharePoint REST API proving the token does not work calling the SharePoint REST API method
            try
            {
                Console.WriteLine($"Calling {sharepointTenantUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties");
                string sharepointResponse = await MakeSharePointRestApiCall($"{sharepointTenantUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties", sharePointAuthResult.AccessToken);
                Console.WriteLine($"[SharePoint REST API response] (first 500 chars):\n{sharepointResponse.Substring(0, 500)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"With just the SharePoint scope (allsites.manage) calling the SharePoint API user profile service we get this error because we don't have enough permissions:\n{ex.Message}");
            }
        }

        private static async Task<AuthenticationResult> AuthenticateWithAzureADAsync(IEnumerable<string> scopes, IAccount account)
        {
            if (account == null)
            {
                return await publicClientApp
                    .AcquireTokenInteractive(scopes)
                    .WithAuthority($"https://login.microsoftonline.com/organizations")
                    .ExecuteAsync();
            }
            else
            {
                try
                {
                    return await publicClientApp
                        .AcquireTokenSilent(scopes, account)
                        .WithAuthority($"https://login.microsoftonline.com/organizations")
                        .ExecuteAsync();
                }
                catch (MsalUiRequiredException ex)
                {
                    return await publicClientApp
                        .AcquireTokenInteractive(scopes)
                        .WithAccount(account)
                        .ExecuteAsync();
                }
            }
        }

        private static async Task<string> MakeSharePointRestApiCall(string uri, string accessToken)
        {
            HttpClient sharepointHttpClient = new HttpClient();
            sharepointHttpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            return await sharepointHttpClient.GetStringAsync(uri);
        }
    }
}