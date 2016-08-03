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
    class GroupSnippets
    {
        const string serviceEndpoint = "https://graph.microsoft.com/v1.0/";

        // Returns all of the groups in your tenant's directory.
        public static async Task<List<string>> GetGroupsAsync()
        {
            var groups = new List<string>();
            JObject jResult = null;

            try
            {
                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenForUserAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                // Endpoint for all groups in the tenant
                Uri groupsEndpoint = new Uri(serviceEndpoint + "myOrganization/groups");

                HttpResponseMessage response = await client.GetAsync(groupsEndpoint);

                if (response.IsSuccessStatusCode)
                {
                    string responseContent = await response.Content.ReadAsStringAsync();
                    jResult = JObject.Parse(responseContent);

                    foreach (JObject group in jResult["value"])
                    {
                        string groupId = (string)group["id"];
                        groups.Add(groupId);
                        Debug.WriteLine("Got group: " + groupId);
                    }
                }

                else
                {
                    Debug.WriteLine("We could not get groups. The request returned this status code: " + response.StatusCode);
                    return null;
                }

                return groups;

            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not get groups: " + e.Message);
                return null;
            }

        }


        // Returns the display name of a specific group.
        public static async Task<string> GetGroupAsync(string groupId)
        {
            string groupName = null;
            JObject jResult = null;

            try
            {
                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenForUserAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                // Endpoint for the specified group
                Uri groupEndpoint = new Uri(serviceEndpoint + "myOrganization/groups/" + groupId);

                HttpResponseMessage response = await client.GetAsync(groupEndpoint);

                if (response.IsSuccessStatusCode)
                {
                    string responseContent = await response.Content.ReadAsStringAsync();
                    jResult = JObject.Parse(responseContent);
                    groupName = (string)jResult["displayName"];
                    Debug.WriteLine("Got group: " + groupName);
                }

                else
                {
                    Debug.WriteLine("We could not get the specified group. The request returned this status code: " + response.StatusCode);
                    return null;
                }

            }


            catch (Exception e)
            {
                Debug.WriteLine("We could not get the specified group: " + e.Message);
                return null;

            }

            return groupName;
        }

        public static async Task<List<string>> GetGroupMembersAsync(string groupId)
        {
            var members = new List<string>();
            JObject jResult = null;

            try
            {
                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenForUserAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                // Endpoint for all members of the specified group
                Uri membersEndpoint = new Uri(serviceEndpoint + "myOrganization/groups/" + groupId + "/members");

                HttpResponseMessage response = await client.GetAsync(membersEndpoint);

                if (response.IsSuccessStatusCode)
                {
                    string responseContent = await response.Content.ReadAsStringAsync();
                    jResult = JObject.Parse(responseContent);

                    foreach (JObject member in jResult["value"])
                    {
                        string userName = (string)member["displayName"];
                        members.Add(userName);
                        Debug.WriteLine("Got member: " + userName);
                    }

                }

                else
                {
                    Debug.WriteLine("We could not get the group members. The request returned this status code: " + response.StatusCode);
                    return null;

                }
            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not get the group members: " + e.Message);
                return null;
            }

            return members;

        }

        public static async Task<List<string>> GetGroupOwnersAsync(string groupId)
        {
            var owners = new List<string>();
            JObject jResult = null;

            try
            {
                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenForUserAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                // Endpoint for all owners of the specified group
                Uri ownersEndpoint = new Uri(serviceEndpoint + "myOrganization/groups/" + groupId + "/owners");

                HttpResponseMessage response = await client.GetAsync(ownersEndpoint);

                if (response.IsSuccessStatusCode)
                {
                    string responseContent = await response.Content.ReadAsStringAsync();
                    jResult = JObject.Parse(responseContent);

                    foreach (JObject member in jResult["value"])
                    {
                        string userName = (string)member["displayName"];
                        owners.Add(userName);
                        Debug.WriteLine("Got owner: " + userName);
                    }

                }

                else
                {
                    Debug.WriteLine("We could not get the group owners. The request returned this status code: " + response.StatusCode);
                    return null;

                }
            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not get the group owners: " + e.Message);
                return null;
            }

            return owners;

        }


        // Creates a new security group in the tenant.
        public static async Task<string> CreateGroupAsync(string groupName)
        {
            JObject jResult = null;
            string createdGroupId = null;
            try
            {
                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenForUserAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                // Endpoint for all groups in an organization
                Uri groupsEndpoint = new Uri(serviceEndpoint + "myOrganization/groups");

                // Build contents of post body and convert to StringContent object.
                // Using line breaks for readability.

                string postBody = "{'mailEnabled':false," // Must be false, because only pure security groups can be created with the unified API.
                                + "'displayName':'Group " + groupName + "',"
                                + "'mailNickName':'" + groupName + "',"
                                + "'securityEnabled':true" // Must be true, because only pure security groups can be created with the unified API.
                                + "}";

                var createBody = new StringContent(postBody, System.Text.Encoding.UTF8, "application/json");

                HttpResponseMessage response = await client.PostAsync(groupsEndpoint, createBody);

                if (response.IsSuccessStatusCode)
                {
                    string responseContent = await response.Content.ReadAsStringAsync();
                    jResult = JObject.Parse(responseContent);
                    createdGroupId = (string)jResult["id"];
                    Debug.WriteLine("Created group: " + createdGroupId);
                }

                else
                {
                    Debug.WriteLine("We could not create a group. The request returned this status code: " + response.StatusCode);
                    return null;
                }

            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not create a group: " + e.Message);
                return null;
            }

            return createdGroupId;

        }


        // Updates the description of an existing group.
        public static async Task<bool> UpdateGroupAsync(string groupId)
        {
            bool groupUpdated = false;

            try
            {
                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenForUserAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                // Endpoint for all the specified group.
                Uri groupEndpoint = new Uri(serviceEndpoint + "myOrganization/groups/" + groupId);

                string updateBody = "{ 'Description': 'This is an updated group group.' }";
                var patchBody = new StringContent(updateBody, System.Text.Encoding.UTF8, "application/json");

                // Construct HTTP PATCH request

                var method = new HttpMethod("PATCH");
                var request = new HttpRequestMessage(method, groupEndpoint) { Content = patchBody };

                HttpResponseMessage response = await client.SendAsync(request);

                if (response.IsSuccessStatusCode)
                {
                    groupUpdated = true;
                    Debug.WriteLine("Updated group: " + groupId);
                }

                else
                {
                    Debug.WriteLine("We could not update the group. The request returned this status code: " + response.StatusCode);
                    groupUpdated = false;
                }

            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not update the group: " + e.Message);
                groupUpdated = false;
            }

            return groupUpdated;

        }


        // Deletes an existing group in the tenant.
        public static async Task<bool> DeleteGroupAsync(string groupId)
        {
            bool eventDeleted = false;

            try
            {
                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenForUserAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                // Endpoint for the specified group
                Uri eventEndpoint = new Uri(serviceEndpoint + "myOrganization/groups/" + groupId);

                HttpResponseMessage response = await client.DeleteAsync(eventEndpoint);

                if (response.IsSuccessStatusCode)
                {
                    eventDeleted = true;
                    Debug.WriteLine("Deleted group: " + groupId);
                }

                else
                {
                    Debug.WriteLine("We could not delete the group. The request returned this status code: " + response.StatusCode);
                    eventDeleted = false;
                }

            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not delete the group: " + e.Message);
                eventDeleted = false;
            }

            return eventDeleted;

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
// ""Software""), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:

// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.

// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
// 
//********************************************************* 