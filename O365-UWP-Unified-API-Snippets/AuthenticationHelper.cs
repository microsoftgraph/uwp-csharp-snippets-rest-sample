// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.


using System;
using System.Diagnostics;
using System.Net.Http;
using System.Linq;
using System.Threading.Tasks;
using Windows.Security.Authentication.Web;
using Windows.Security.Authentication.Web.Core;
using Windows.Security.Credentials;
using Windows.Storage;

namespace O365_UWP_Unified_API_Snippets
{
    internal static class AuthenticationHelper
    {
        // The Client ID is used by the application to uniquely identify itself to Microsoft Azure Active Directory (AD).
        static string clientId = App.Current.Resources["ida:ClientID"].ToString();

        // You'll create your tenant-specific authority from the tenant domain and AADInstance URI
        static string tenant = App.Current.Resources["ida:Domain"].ToString();
        static string AADInstance = App.Current.Resources["ida:AADInstance"].ToString();

        //Use the domain-specific authority when you're authenticating users from a single tenant only.
        static string authority = AADInstance + tenant;

        // Use "organizations" as your authority when you want the app to work on any Azure Tenant.
        // static string authority = "organizations";

        // To authenticate to the unified API, the client needs to know its App ID URI.
        public const string ResourceUrl = "https://graph.microsoft.com/";
        private const string provider = "https://login.windows.net";

        private static WebAccountProvider aadAccountProvider = null;
        private static WebAccount userAccount = null;

        // Store account-specific settings so that the app can remember that a user has already signed in.
        public static ApplicationDataContainer _settings = ApplicationData.Current.RoamingSettings;



        // Get an access token for the given context and resourceId. An attempt is first made to 
        // acquire the token silently. If that fails, then we try to acquire the token by prompting the user.
        public static async Task<string> GetTokenHelperAsync()
        {

            string token = null;

            aadAccountProvider = await WebAuthenticationCoreManager.FindAccountProviderAsync(provider, authority);

            // Check if there's a record of the last account used with the app.
            var userID = _settings.Values["userID"];

            if (userID != null)
            {

                WebTokenRequest webTokenRequest = new WebTokenRequest(aadAccountProvider, string.Empty, clientId);
                webTokenRequest.Properties.Add("resource", ResourceUrl);
                webTokenRequest.Properties.Add("authority", provider);

                // Get an account object for the user.
                userAccount = await WebAuthenticationCoreManager.FindAccountAsync(aadAccountProvider, (string)userID);


                // Ensure that the saved account works for getting the token we need.
                WebTokenRequestResult webTokenRequestResult = await WebAuthenticationCoreManager.RequestTokenAsync(webTokenRequest, userAccount);
                if (webTokenRequestResult.ResponseStatus == WebTokenRequestStatus.Success || webTokenRequestResult.ResponseStatus == WebTokenRequestStatus.AccountSwitch)
                {
                    WebTokenResponse webTokenResponse = webTokenRequestResult.ResponseData[0];
                    userAccount = webTokenResponse.WebAccount;
                    token = webTokenResponse.Token;

                }
                else
                {
                    // The saved account could not be used for getting a token.
                    // Make sure that the UX is ready for a new sign in.
                    SignOut();
                }

            }
            else
            {
                // There is no recorded user. Start a sign-in flow without imposing a specific account.

                WebTokenRequest webTokenRequest = new WebTokenRequest(aadAccountProvider, string.Empty, clientId, WebTokenRequestPromptType.ForceAuthentication);
                webTokenRequest.Properties.Add("resource", ResourceUrl);
                webTokenRequest.Properties.Add("authority", provider);
                WebTokenRequestResult webTokenRequestResult = await WebAuthenticationCoreManager.RequestTokenAsync(webTokenRequest);
                if (webTokenRequestResult.ResponseStatus == WebTokenRequestStatus.Success)
                {
                    WebTokenResponse webTokenResponse = webTokenRequestResult.ResponseData[0];
                    userAccount = webTokenResponse.WebAccount;
                    token = webTokenResponse.Token;

                }
            }

            // We succeeded in getting a valid user.
            if (userAccount != null)
            {
                // Save user ID in local storage.
                _settings.Values["userID"] = userAccount.Id;
                _settings.Values["userEmail"] = userAccount.UserName;
                _settings.Values["userName"] = userAccount.Properties["DisplayName"];

                return token;
            }

            // We didn't succeed in getting a valid user. Clear the app settings so that another user can sign in.
            else
            {

                SignOut();
                return null;
            }


        }

        /// <summary>
        /// Signs the user out of the service.
        /// </summary>
        public static void SignOut()
        {
            //Clear stored values from last authentication.
            _settings.Values["userID"] = null;
            _settings.Values["userEmail"] = null;
            _settings.Values["userName"] = null;

        }

        public static string GetAppRedirectURI()
        {
            // Windows 10 universal apps require redirect URI in the format below. Add a breakpoint to this line, and run the app before you register it so that
            // you can supply the correct redirect URI value.
            return string.Format("ms-appx-web://microsoft.aad.brokerplugin/{0}", WebAuthenticationBroker.GetCurrentApplicationCallbackUri().Host).ToUpper();
        }

    }
}

//********************************************************* 
// 
//O365-UWP-Unified-API-Snippets, https://github.com/OfficeDev/O365-UWP-Unified-API-Snippets
//
//Copyright (c) Microsoft Corporation
//All rights reserved. 
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:

// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.

// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
// 
//********************************************************* 