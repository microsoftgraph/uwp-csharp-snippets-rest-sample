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
using Microsoft.Identity.Client;

namespace O365_UWP_Unified_API_Snippets
{
    public class AuthenticationHelper
    {
        // The Client ID is used by the application to uniquely identify itself to the v2.0 authentication endpoint.
        static string clientId = App.Current.Resources["ida:ClientID"].ToString();

        public static PublicClientApplication IdentityClientApp = new PublicClientApplication(clientId)
        {
            RedirectUri = App.Current.Resources["ida:ReturnUrl"].ToString()
        };

        public static string TokenForUser = null;
        public static DateTimeOffset Expiration;
        public static ApplicationDataContainer _settings = ApplicationData.Current.RoamingSettings;

        /// <summary>
        /// Get Token for User.
        /// </summary>
        /// <returns>Token for user.</returns>
        public static async Task<string> GetTokenForUserAsync()
        {
            AuthenticationResult authResult;
            var scopes = new string[]
   
             {
                "https://graph.microsoft.com/User.Read",
                "https://graph.microsoft.com/User.ReadWrite",
                "https://graph.microsoft.com/User.ReadBasic.All",
                "https://graph.microsoft.com/Mail.Send",
                "https://graph.microsoft.com/Calendars.ReadWrite",
                "https://graph.microsoft.com/Mail.ReadWrite",
                "https://graph.microsoft.com/Files.ReadWrite",

                 // Admin-only scopes. Uncomment these if you're running the sample with an admin work account.
                 // You won't be able to sign in with a non-admin work account if you request these scopes.
                 // These scopes will be ignored if you leave them uncommented and run the sample with a consumer account.
                 // See the MainPage.xaml.cs file for all of the operations that won't work if you're not running the 
                 // sample with an admin work account.
                 //"https://graph.microsoft.com/User.ReadWrite.All",
                 //"https://graph.microsoft.com/Group.ReadWrite.All",
                 //"https://graph.microsoft.com/Directory.AccessAsUser.All"
             };

            try
            {
                authResult = await IdentityClientApp.AcquireTokenSilentAsync(scopes, IdentityClientApp.Users.FirstOrDefault());
                TokenForUser = authResult.AccessToken;
                // save user ID in local storage
                _settings.Values["userEmail"] = authResult.User.DisplayableId;
                _settings.Values["userName"] = authResult.User.Name;
            }
            catch (MsalUiRequiredException)
            {
                authResult = await IdentityClientApp.AcquireTokenAsync(scopes);

                TokenForUser = authResult.AccessToken;
                Expiration = authResult.ExpiresOn;

                // save user ID in local storage
                _settings.Values["userEmail"] = authResult.User.DisplayableId;
                _settings.Values["userName"] = authResult.User.Name;
            }
            catch (Exception exc)
            {
                Debug.WriteLine(exc.StackTrace);
            }

            return TokenForUser;
        }

        /// <summary>
        /// Signs the user out of the service.
        /// </summary>
        public static void SignOut()
        {
            foreach (var user in IdentityClientApp.Users)
            {
                IdentityClientApp.Remove(user);
            }

            TokenForUser = null;

            //Clear stored values from last authentication.
            _settings.Values["userEmail"] = null;
            _settings.Values["userName"] = null;

        }

    }
}

//********************************************************* 
// 
//O365-UWP-Microsoft-Graph-Snippets, https://github.com/OfficeDev/O365-UWP-Microsoft-Graph-Snippets
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