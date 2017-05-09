// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Diagnostics;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace O365_UWP_Unified_API_Snippets
{
    class BatchingSnippets
    {
        const string Endpoint = "https://graph.microsoft.com/beta/$batch";

        internal static async Task<string> ParallelBatchCallAsync()
        {
            var token = await AuthenticationHelper.GetTokenForUserAsync();

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Post, Endpoint))
                {
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    JArray batchRequests = new JArray
                        {
                            new JObject
                            {
                                { "url", "/me" },
                                { "method", "GET" },
                                { "id", "1" }
                            },
                            new JObject
                            {
                                { "url", "/me/manager" },
                                { "method", "GET" },
                                { "id", "2" }
                            },
                            new JObject
                            {
                                { "url", "me/messages?$top=5" },
                                { "method", "GET" },
                                { "id", "3" }
                            },
                            new JObject
                            {
                                { "url", "/me/photo/$value" },
                                { "method", "GET" },
                                { "id", "4" }
                            },
                        };

                    JObject batchPayload = new JObject(new JProperty("requests", batchRequests));

                    request.Content = new StringContent(JsonConvert.SerializeObject(batchPayload), Encoding.UTF8, "application/json");
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
                            Debug.WriteLine("We could not run parallel batch request. The request returned this status code: " + response.StatusCode);
                            return null;
                        }
                    }
                }
            }
        }

        internal static async Task<string> SequentialBatchCallAsync()
        {
            var token = await AuthenticationHelper.GetTokenForUserAsync();
            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Post, Endpoint))
                {
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    JArray batchRequests = new JArray
                    {
                        new JObject
                        {
                            { "url", "/me/events?$top=2" },
                            { "method", "GET" },
                            { "id", "1" }
                        },
                        new JObject
                        {
                            { "url", "/me/drive/root/children" },
                            { "method", "POST" },
                            { "id", "2" },
                            { "body", new JObject
                                    {
                                        { "name", "BatchingTestFolder"},
                                        { "folder", new JObject()}
                                    }
                            },
                            { "headers", new JObject { { "Content-Type", "application/json" } } },
                            { "dependsOn", new JArray("1")}
                        },
                        new JObject
                        {
                            { "url", "/me/drive/root/children/BatchingTestFolder" },
                            { "method", "GET" },
                            { "id", "3" },
                            { "dependsOn", new JArray("2")}
                        },
                        new JObject
                        {
                            { "url", "/me/drive/root/children/BatchingTestFolder" },
                            { "method", "DELETE" },
                            { "id", "4" },
                            { "dependsOn", new JArray("3")}
                        }
                    };

                    JObject batchPayload = new JObject(new JProperty("requests", batchRequests));

                    request.Content = new StringContent(JsonConvert.SerializeObject(batchPayload), Encoding.UTF8, "application/json");
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
                            Debug.WriteLine("We could not run sequential batch request. The request returned this status code: " + response.StatusCode);
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