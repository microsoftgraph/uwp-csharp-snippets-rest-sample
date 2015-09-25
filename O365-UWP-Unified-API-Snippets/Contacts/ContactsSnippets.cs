// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace O365_UWP_Unified_API_Snippets
{
    class ContactsSnippets
    {
        const string serviceEndpoint = "https://graph.microsoft.com/beta/";

        // Returns all of the contacts in your tenant's directory.
        public static async Task<List<string>> GetContactsAsync()
        {
            var contacts = new List<string>();
            JObject jResult = null;

            try
            {
                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenHelperAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                // Endpoint for all contacts in your organization
                Uri contactsEndpoint = new Uri(serviceEndpoint + "myOrganization/contacts");

                HttpResponseMessage response = await client.GetAsync(contactsEndpoint);

                if (response.IsSuccessStatusCode)
                {
                    string responseContent = await response.Content.ReadAsStringAsync();
                    jResult = JObject.Parse(responseContent);

                    foreach (JObject contact in jResult["value"])
                    {
                        string contactName = (string)contact["displayName"];
                        contacts.Add(contactName);
                        Debug.WriteLine("Got contact: " + contactName);
                    }
                }

                else
                {
                    Debug.WriteLine("We could not get contacts. The request returned this status code: " + response.StatusCode);
                    return null;
                }

                return contacts;

            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not get contacts: " + e.Message);
                return null;
            }
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