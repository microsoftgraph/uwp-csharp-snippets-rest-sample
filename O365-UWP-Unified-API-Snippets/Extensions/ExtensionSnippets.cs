// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace O365_UWP_Unified_API_Snippets
{
    class ExtensionSnippets
    {
        const string serviceEndpoint = "https://graph.microsoft.com/beta";
        const string domainName = "adatumisv";

        internal static async Task<string> GetOpenExtensionsAsync()
        {
            string endpoint = serviceEndpoint + "/me/extensions/sampleSettings";

            var token = await AuthenticationHelper.GetTokenForUserAsync();

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, endpoint))
                {
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                    // This header has been added to identify our sample in the Microsoft Graph service. If extracting this code for your project please remove.
                    request.Headers.Add("SampleID", "aspnet-connect-rest-sample");
                    using (HttpResponseMessage response = await client.SendAsync(request))
                    {
                        if (response.StatusCode == HttpStatusCode.OK)
                        {
                            var json = JObject.Parse(await response.Content.ReadAsStringAsync());
                            return json.ToString();
                        }
                        else
                        {
                            Debug.WriteLine("We could not get extensions. The request returned this status code: " + response.StatusCode);
                            return null;
                        }
                    }
                }
            }
        }

        internal static async Task<string> SetOpenExtensionsAsync()
        {
            string endpoint = null;
            string openExtensionPropertyName = "prop1";
            string openExtensionPropertyValue = "value1";
            string settingName = "sampleSettings";
            JObject setting = new JObject
            {
                { "id", settingName },
                {openExtensionPropertyName, openExtensionPropertyValue}
            };

            endpoint = serviceEndpoint + "/me/extensions/sampleSettings";
            var token = await AuthenticationHelper.GetTokenForUserAsync();

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(new HttpMethod("PATCH"), endpoint))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                    // This header has been added to identify our sample in the Microsoft Graph service. If extracting this code for your project please remove.
                    request.Headers.Add("SampleID", "aspnet-connect-rest-sample");
                    request.Content = new StringContent(JsonConvert.SerializeObject(setting), Encoding.UTF8, "application/json");
                    using (HttpResponseMessage response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            return "Success!";
                        }
                        else
                        {
                            endpoint = serviceEndpoint + "/me/extensions";
                            using (var request2 = new HttpRequestMessage(new HttpMethod("POST"), endpoint))
                            {
                                request2.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                                request2.Headers.Add("SampleID", "aspnet-connect-rest-sample");
                                request2.Content = new StringContent(JsonConvert.SerializeObject(setting), Encoding.UTF8, "application/json");

                                using (HttpResponseMessage response2 = await client.SendAsync(request2))
                                {
                                    if (response2.IsSuccessStatusCode)
                                    {
                                        return "Success!";
                                    }

                                    return null;
                                }
                            }
                        }
                    }
                }
            }
        }

        internal static async Task<string> DeleteSchemaExtensionsAsync()
        {
            string extName = domainName + "_hrprofile";
            string endpoint = serviceEndpoint + "/schemaExtensions/" + extName;
            var token = await AuthenticationHelper.GetTokenForUserAsync();

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Delete, endpoint))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                    // This header has been added to identify our sample in the Microsoft Graph service. If extracting this code for your project please remove.
                    request.Headers.Add("SampleID", "aspnet-connect-rest-sample");
                    using (HttpResponseMessage response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            return "Success!";
                        }
                        else
                        {
                            return null;
                        }
                    }
                }
            }
        }

        internal static async Task<string> GetSchemaExtensionValueAsync()
        {
            string extName = domainName + "_hrprofile";
            string endpoint = serviceEndpoint + "/me?$select=" + extName;
            var token = await AuthenticationHelper.GetTokenForUserAsync();

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, endpoint))
                {
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                    // This header has been added to identify our sample in the Microsoft Graph service. If extracting this code for your project please remove.
                    request.Headers.Add("SampleID", "aspnet-connect-rest-sample");
                    using (HttpResponseMessage response = await client.SendAsync(request))
                    {
                        if (response.StatusCode == HttpStatusCode.OK)
                        {
                            var json = JObject.Parse(await response.Content.ReadAsStringAsync());
                            return json.ToString();
                        }

                        return null;
                    }
                }
            }
        }

        internal static async Task<string> SetExtensionValueAsync()
        {
            string endpoint = serviceEndpoint + "/me";
            var token = await AuthenticationHelper.GetTokenForUserAsync();
            string extName = domainName + "_hrprofile";
            string propName = "p1";
            string propValue = "value";

            JObject extensionValue = new JObject
            {
                {
                    extName, new JObject {
                        {propName, propValue }
                    }
                },
            };

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(new HttpMethod("PATCH"), endpoint))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                    // This header has been added to identify our sample in the Microsoft Graph service. If extracting this code for your project please remove.
                    request.Headers.Add("SampleID", "aspnet-connect-rest-sample");
                    request.Content = new StringContent(JsonConvert.SerializeObject(extensionValue), Encoding.UTF8, "application/json");
                    using (HttpResponseMessage response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            return "Success";
                        }
                        else
                        {
                            return null;
                        }
                    }
                }
            }
        }

        internal static async Task<string> GetSchemaExtensionsAsync()
        {
            string extName = domainName + "_hrprofile";
            string endpoint = serviceEndpoint + "/schemaExtensions/" + extName;
            var token = await AuthenticationHelper.GetTokenForUserAsync();

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, endpoint))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                    // This header has been added to identify our sample in the Microsoft Graph service. If extracting this code for your project please remove.
                    request.Headers.Add("SampleID", "aspnet-connect-rest-sample");
                    using (HttpResponseMessage response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            return "Success!";
                        }
                        else
                        {
                            return null;
                        }
                    }
                }
            }
        }

        internal static async Task<string> RegisterSchemaExtensionsAsync()
        {
            string endpoint = serviceEndpoint + "/schemaExtensions";
            string extName = domainName + "_hrprofile";
            string propName = "p1";
            var token = await AuthenticationHelper.GetTokenForUserAsync();

            JObject schemaExt = new JObject
            {
                { "status", "InDevelopment"},
                { "id", extName },
                { "targetTypes",new JArray(new string[] { "User"} )},
                { "description","Extension description"},
                { "properties", new JArray(new JObject[]
                    {
                        new JObject { { "name", propName }, { "type", "String" } }
                    })
                }
            };

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);

                    // This header has been added to identify our sample in the Microsoft Graph service. If extracting this code for your project please remove.
                    request.Headers.Add("SampleID", "aspnet-connect-rest-sample");
                    request.Content = new StringContent(JsonConvert.SerializeObject(schemaExt), Encoding.UTF8, "application/json");
                    using (HttpResponseMessage response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            return "Success!";
                        }
                        else
                        {
                            var json = JObject.Parse(await response.Content.ReadAsStringAsync());
                            return null;
                        }
                    }
                }
            }
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