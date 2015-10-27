// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace O365_UWP_Unified_API_Snippets
{
    class UserSnippets
    {
        const string serviceEndpoint = "https://graph.microsoft.com/v1.0/";
        static string tenant = App.Current.Resources["ida:Domain"].ToString();

        // Returns information about the signed-in user from Azure Active Directory.
        public static async Task<string> GetMeAsync()
        {
            string currentUser = null;
            JObject jResult = null;

            try
            {
                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenHelperAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                // Endpoint for all users in an organization
                Uri usersEndpoint = new Uri(serviceEndpoint + "me");

                HttpResponseMessage response = await client.GetAsync(usersEndpoint);

                if (response.IsSuccessStatusCode)
                {
                    string responseContent = await response.Content.ReadAsStringAsync();
                    jResult = JObject.Parse(responseContent);
                    currentUser = (string)jResult["displayName"];
                    Debug.WriteLine("Got user: " + currentUser);
                }

                else
                {
                    Debug.WriteLine("We could not get the current user. The request returned this status code: " + response.StatusCode);
                    return null;
                }

            }


            catch (Exception e)
            {
                Debug.WriteLine("We could not get the current user: " + e.Message);
                return null;

            }

            return currentUser;
        }


        // Returns all of the users in the directory of the signed-in user's tenant. 
        public static async Task<List<string>> GetUsersAsync()
        {
            var users = new List<string>();
            JObject jResult = null;

            try
            {
                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenHelperAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                // Endpoint for all users in an organization
                Uri usersEndpoint = new Uri(serviceEndpoint + "myOrganization/users");

                HttpResponseMessage response = await client.GetAsync(usersEndpoint);

                if (response.IsSuccessStatusCode)
                {
                    string responseContent = await response.Content.ReadAsStringAsync();
                    jResult = JObject.Parse(responseContent);

                    foreach (JObject user in jResult["value"])
                    {
                        string userName = (string)user["displayName"];
                        users.Add(userName);
                        Debug.WriteLine("Got user: " + userName);
                    }
                }

                else
                {
                    Debug.WriteLine("We could not get users. The request returned this status code: " + response.StatusCode);
                    return null;
                }

                return users;

            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not get users: " + e.Message);
                return null;
            }


        }

        // Creates a new user in the signed-in user's tenant. This snippet requires an admin account.
        public static async Task<string> CreateUserAsync(string userName)
        {
            JObject jResult = null;
            string createdUserName = null;
            try
            {
                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenHelperAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                // Endpoint for all users in an organization
                Uri usersEndpoint = new Uri(serviceEndpoint + "myOrganization/users");

                // Build contents of post body and convert to StringContent object.
                // Using line breaks for readability.

                string postBody = "{'accountEnabled':true,"
                                + "'displayName':'User " + userName + "',"
                                + "'mailNickName':'" + userName + "',"
                                + "'passwordProfile': {'password': 'pass@word1','forceChangePasswordNextLogin': false },"
                                + "'userPrincipalName':'" + userName + "@" + tenant + "'}";

                var createBody = new StringContent(postBody, System.Text.Encoding.UTF8, "application/json");

                HttpResponseMessage response = await client.PostAsync(usersEndpoint, createBody);

                if (response.IsSuccessStatusCode)
                {
                    string responseContent = await response.Content.ReadAsStringAsync();
                    jResult = JObject.Parse(responseContent);
                    createdUserName = (string)jResult["displayName"];
                    Debug.WriteLine("Created user: " + createdUserName);
                }

                else
                {
                    Debug.WriteLine("We could not create a user. The request returned this status code: " + response.StatusCode);
                    return null;
                }

            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not create a user: " + e.Message);
                return null;
            }

            return createdUserName;

        }

        // Gets the signed-in user's drive.
        public static async Task<string> GetCurrentUserDriveAsync()
        {
            string currentUserDriveId = null;
            JObject jResult = null;

            try
            {
                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenHelperAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                // Endpoint for the current user's drive
                Uri usersEndpoint = new Uri(serviceEndpoint + "me/drive");

                HttpResponseMessage response = await client.GetAsync(usersEndpoint);

                if (response.IsSuccessStatusCode)
                {
                    string responseContent = await response.Content.ReadAsStringAsync();
                    jResult = JObject.Parse(responseContent);
                    currentUserDriveId = (string)jResult["id"];
                    Debug.WriteLine("Got user drive: " + currentUserDriveId);
                }

                else
                {
                    Debug.WriteLine("We could not get the current user drive. The request returned this status code: " + response.StatusCode);
                    return null;
                }

            }


            catch (Exception e)
            {
                Debug.WriteLine("We could not get the current user drive: " + e.Message);
                return null;

            }

            return currentUserDriveId;

        }

        // Gets the signed-in user's calendar events.

        public static async Task<List<string>> GetEventsAsync()
        {
            var events = new List<string>();
            JObject jResult = null;

            try
            {
                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenHelperAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                // Endpoint for the current user's events
                Uri usersEndpoint = new Uri(serviceEndpoint + "me/events?$select=id");

                HttpResponseMessage response = await client.GetAsync(usersEndpoint);

                if (response.IsSuccessStatusCode)
                {
                    string responseContent = await response.Content.ReadAsStringAsync();
                    jResult = JObject.Parse(responseContent);

                    foreach (JObject calendarEvent in jResult["value"])
                    {
                        string eventId = (string)calendarEvent["Id"];
                        events.Add(eventId);
                        Debug.WriteLine("Got event: " + eventId);
                    }

                }

                else
                {
                    Debug.WriteLine("We could not get the current user's events. The request returned this status code: " + response.StatusCode);
                    return null;

                }
            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not get the current user's events: " + e.Message);
                return null;
            }

            return events;
        }

        // Creates a new event in the signed-in user's tenant.
        // Important note: This will create a user with a weak password. Consider deleting this user after you run the sample.
        public static async Task<string> CreateEventAsync()
        {
            JObject jResult = null;
            string createdEventId = null;
            try
            {
                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenHelperAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                // Endpoint for the current user's events
                Uri eventsEndpoint = new Uri(serviceEndpoint + "me/events");

                // Build contents of post body and convert to StringContent object.
                // Using line breaks for readability.

                // Specifying the round-trip format specifier ("o") to the DateTimeOffset.ToString() method
                // so that the datetime string can be converted into an Edm.DateTimeOffset object:
                // https://msdn.microsoft.com/en-us/library/az4se3k1(v=vs.110).aspx#Roundtrip

                string postBody = "{'Subject':'Weekly Sync',"
                                + "'Location':{'DisplayName':'Water Cooler'},"
                                + "'Attendees':[{'Type':'Required','EmailAddress': {'Address':'mara@fabrikam.com'} }],"
                                + "'Start':'" + new DateTimeOffset(new DateTime(2014, 12, 1, 9, 30, 0)).ToString("o") + "',"
                                + "'End':'" + new DateTimeOffset(new DateTime(2014, 12, 1, 10, 0, 0)).ToString("o") + "',"
                                + "'Body':{'Content': 'Status updates, blocking issues, and next steps.', 'ContentType':'Text'}}";

                var createBody = new StringContent(postBody, System.Text.Encoding.UTF8, "application/json");

                HttpResponseMessage response = await client.PostAsync(eventsEndpoint, createBody);

                if (response.IsSuccessStatusCode)
                {
                    string responseContent = await response.Content.ReadAsStringAsync();
                    jResult = JObject.Parse(responseContent);
                    createdEventId = (string)jResult["Id"];
                    Debug.WriteLine("Created event: " + createdEventId);
                }

                else
                {
                    Debug.WriteLine("We could not create an event. The request returned this status code: " + response.StatusCode);
                    return null;
                }

            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not create an event: " + e.Message);
                return null;
            }

            return createdEventId;

        }

        // Updates the subject of an existing event in the signed-in user's tenant.
        public static async Task<bool> UpdateEventAsync(string eventId)
        {
            bool eventUpdated = false;

            try
            {
                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenHelperAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                // Endpoint for the specified event.
                Uri eventEndpoint = new Uri(serviceEndpoint + "me/events/" + eventId);

                string updateBody = "{ 'Subject': 'Sync of the Week' }";
                var patchBody = new StringContent(updateBody, System.Text.Encoding.UTF8, "application/json");

                // Construct HTTP PATCH request

                var method = new HttpMethod("PATCH");
                var request = new HttpRequestMessage(method, eventEndpoint) { Content = patchBody };

                HttpResponseMessage response = await client.SendAsync(request);

                if (response.IsSuccessStatusCode)
                {
                    eventUpdated = true;
                    Debug.WriteLine("Updated event: " + eventId);
                }

                else
                {
                    Debug.WriteLine("We could not update the event. The request returned this status code: " + response.StatusCode);
                    eventUpdated = false;
                }

            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not update the event: " + e.Message);
                eventUpdated = false;
            }

            return eventUpdated;

        }

        // Deletes an existing event in the signed-in user's tenant.
        public static async Task<bool> DeleteEventAsync(string eventId)
        {
            bool eventDeleted = false;

            try
            {
                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenHelperAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                // Endpoint for the specified event
                Uri eventEndpoint = new Uri(serviceEndpoint + "me/events/" + eventId);

                HttpResponseMessage response = await client.DeleteAsync(eventEndpoint);

                if (response.IsSuccessStatusCode)
                {
                    eventDeleted = true;
                    Debug.WriteLine("Deleted event: " + eventId);
                }

                else
                {
                    Debug.WriteLine("We could not delete the event. The request returned this status code: " + response.StatusCode);
                    eventDeleted = false;
                }

            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not delete the event: " + e.Message);
                eventDeleted = false;
            }

            return eventDeleted;

        }

        // Returns the first page of the signed-in user's messages.
        public static async Task<List<string>> GetMessagesAsync()
        {
            var messages = new List<string>();
            JObject jResult = null;

            try
            {
                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenHelperAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                // Endpoint for all messages in the current user's mailbox
                Uri messagesEndpoint = new Uri(serviceEndpoint + "me/messages");

                HttpResponseMessage response = await client.GetAsync(messagesEndpoint);

                if (response.IsSuccessStatusCode)
                {
                    string responseContent = await response.Content.ReadAsStringAsync();
                    jResult = JObject.Parse(responseContent);

                    foreach (JObject user in jResult["value"])
                    {
                        string subject = (string)user["Subject"];
                        messages.Add(subject);
                        Debug.WriteLine("Got message: " + subject);
                    }
                }

                else
                {
                    Debug.WriteLine("We could not get messages. The request returned this status code: " + response.StatusCode);
                    return null;
                }

                return messages;

            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not get messages: " + e.Message);
                return null;
            }


        }

        // Updates the subject of an existing event in the signed-in user's tenant.
        public static async Task<bool> SendMessageAsync(
            string Subject,
            string Body,
            string RecipientAddress
            )
        {
            bool emailSent = false;

            try
            {

                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenHelperAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                // Endpoint for sending mail from the current user's mailbox
                Uri messageEndpoint = new Uri(serviceEndpoint + "me/SendMail");


                string recipientJSON = "{'EmailAddress':{'Address':'" + RecipientAddress + "'}}";

                // Build contents of post body and convert to StringContent object.
                // Using line breaks for readability.
                string postBody = "{'Message':{"
                    + "'Body':{ "
                    + "'Content': '" + Body + "',"
                    + "'ContentType':'HTML'},"
                    + "'Subject':'" + Subject + "',"
                    + "'ToRecipients':[" + recipientJSON + "]},"
                    + "'SaveToSentItems':true}";

                var emailBody = new StringContent(postBody, System.Text.Encoding.UTF8, "application/json");

                HttpResponseMessage response = await client.PostAsync(messageEndpoint, emailBody);

                if (!response.IsSuccessStatusCode)
                {

                    Debug.WriteLine("We could not send the message. The request returned this status code: " + response.StatusCode);
                    emailSent = false;
                }
                else
                {
                    emailSent = true;
                }


            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not send the message. The request returned this status code: " + e.Message);
                emailSent = false;
            }

            return emailSent;
        }

        // Gets the signed-in user's manager.
        public static async Task<string> GetCurrentUserManagerAsync()
        {
            string currentUserManager = null;
            JObject jResult = null;

            try
            {
                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenHelperAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                // Endpoint for the current user's manager
                Uri managerEndpoint = new Uri(serviceEndpoint + "me/manager");

                HttpResponseMessage response = await client.GetAsync(managerEndpoint);

                if (response.IsSuccessStatusCode)
                {
                    string responseContent = await response.Content.ReadAsStringAsync();
                    jResult = JObject.Parse(responseContent);
                    currentUserManager = (string)jResult["displayName"];
                    Debug.WriteLine("Got manager: " + currentUserManager);
                }

                else
                {
                    Debug.WriteLine("We could not get the current user's manager. The request returned this status code: " + response.StatusCode);
                    return null;
                }

            }


            catch (Exception e)
            {
                Debug.WriteLine("We could not get the current user's manager: " + e.Message);
                return null;

            }

            return currentUserManager;

        }

        // Gets the signed-in user's direct reports.
        public static async Task<List<string>> GetDirectReportsAsync()
        {
            var directReports = new List<string>();
            JObject jResult = null;

            try
            {
                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenHelperAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                // Endpoint for the current user's direct reports
                Uri directsEndpoint = new Uri(serviceEndpoint + "me/directReports");

                HttpResponseMessage response = await client.GetAsync(directsEndpoint);

                if (response.IsSuccessStatusCode)
                {
                    string responseContent = await response.Content.ReadAsStringAsync();
                    jResult = JObject.Parse(responseContent);

                    foreach (JObject user in jResult["value"])
                    {
                        string userName = (string)user["displayName"];
                        directReports.Add(userName);
                        Debug.WriteLine("Got direct report: " + userName);
                    }
                }

                else
                {
                    Debug.WriteLine("We could not get direct reports. The request returned this status code: " + response.StatusCode);
                    return null;
                }

                return directReports;

            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not get direct reports: " + e.Message);
                return null;
            }


        }


        // Gets the signed-in user's photo.
        public static async Task<string> GetCurrentUserPhotoAsync()
        {
            string currentUserPhotoId = null;
            JObject jResult = null;

            try
            {
                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenHelperAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                // Endpoint for the current user's photo
                Uri photoEndpoint = new Uri(serviceEndpoint + "me/photo");

                HttpResponseMessage response = await client.GetAsync(photoEndpoint);

                if (response.IsSuccessStatusCode)
                {
                    string responseContent = await response.Content.ReadAsStringAsync();
                    jResult = JObject.Parse(responseContent);
                    currentUserPhotoId = (string)jResult["Id"];
                    Debug.WriteLine("Got user photo: " + currentUserPhotoId);
                }

                else
                {
                    Debug.WriteLine("We could not get the current user photo. The request returned this status code: " + response.StatusCode);
                    return null;
                }

            }


            catch (Exception e)
            {
                Debug.WriteLine("We could not get the current user photo: " + e.Message);
                return null;

            }

            return currentUserPhotoId;

        }

        // Gets the groups that the signed-in user is a member of.
        public static async Task<List<string>> GetCurrentUserGroupsAsync()
        {
            var memberOfGroups = new List<string>();
            JObject jResult = null;

            try
            {
                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenHelperAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                // Endpoint for the current user's groups
                Uri memberOfEndpoint = new Uri(serviceEndpoint + "me/memberOf");

                HttpResponseMessage response = await client.GetAsync(memberOfEndpoint);

                if (response.IsSuccessStatusCode)
                {
                    string responseContent = await response.Content.ReadAsStringAsync();
                    jResult = JObject.Parse(responseContent);

                    foreach (JObject group in jResult["value"])
                    {
                        string groupId = (string)group["objectId"];
                        memberOfGroups.Add(groupId);
                        Debug.WriteLine("Got group: " + groupId);
                    }
                }

                else
                {
                    Debug.WriteLine("We could not get user groups. The request returned this status code: " + response.StatusCode);
                    return null;
                }

                return memberOfGroups;

            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not get user groups: " + e.Message);
                return null;
            }


        }


        public static async Task<List<string>> GetCurrentUserFilesAsync()
        {
            var files = new List<string>();
            JObject jResult = null;

            try
            {
                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenHelperAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                // Endpoint for all files and folders for a user
                Uri filesEndpoint = new Uri(serviceEndpoint + "me/drive/root/children");

                HttpResponseMessage response = await client.GetAsync(filesEndpoint);

                if (response.IsSuccessStatusCode)
                {
                    string responseContent = await response.Content.ReadAsStringAsync();
                    jResult = JObject.Parse(responseContent);

                    foreach (JObject file in jResult["value"])
                    {
                        string fileName = (string)file["name"];
                        files.Add(fileName);
                        Debug.WriteLine("Got file: " + fileName);
                    }
                }

                else
                {
                    Debug.WriteLine("We could not get user files. The request returned this status code: " + response.StatusCode);
                    return null;
                }

                return files;

            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not get user files: " + e.Message);
                return null;
            }


        }

        // Creates a text file in the user's root directory.
        public static async Task<string> CreateFileAsync(string fileName, string fileContent)
        {
            string createdFileId = null;
            JObject jResult = null;

            try
            {

                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenHelperAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                var fileContentPostBody = new StringContent(fileContent, System.Text.Encoding.UTF8, "text/plain");

                // Endpoint for content in an existing file.
                Uri fileEndpoint = new Uri(serviceEndpoint + "me/drive/root/children/" + fileName + "/content");

                var requestMessage = new HttpRequestMessage(HttpMethod.Put, fileEndpoint)
                {
                    Content = fileContentPostBody
                };


                HttpResponseMessage response = await client.SendAsync(requestMessage);

                if (response.IsSuccessStatusCode)
                {
                    string responseContent = await response.Content.ReadAsStringAsync();
                    jResult = JObject.Parse(responseContent);
                    createdFileId = (string)jResult["id"];
                    Debug.WriteLine("Created file Id: " + createdFileId);


                }
                else
                {
                    Debug.WriteLine("We could not create the file. The request returned this status code: " + response.StatusCode);
                    return null;
                }


            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not create the file. The request returned this status code: " + e.Message);
                return null;
            }

            return createdFileId;
        }

        // Downloads the content of an existing file.
        public static async Task<Stream> DownloadFileAsync(string fileId)
        {
            Stream fileContent = null;

            try
            {
                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenHelperAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                // Endpoint for content in an existing file. Use "/me/drive/root/children/<file name>/content" if you know the name but not the Id.
                Uri fileEndpoint = new Uri(serviceEndpoint + "me/drive/items/" + fileId + "/content");

                HttpResponseMessage response = await client.GetAsync(fileEndpoint);

                if (response.IsSuccessStatusCode)
                {
                    fileContent = await response.Content.ReadAsStreamAsync();
                    Debug.WriteLine("Downloaded file: " + fileId);


                }
                else
                {
                    Debug.WriteLine("We could not download the file. The request returned this status code: " + response.StatusCode);
                    return null;
                }


            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not download the file. The request returned this status code: " + e.Message);
                return null;
            }

            return fileContent;
        }

        // Adds content to a file in the user's root directory.
        public static async Task<bool> UpdateFileAsync(string fileId, string fileContent)
        {
            bool fileUpdated = false;

            try
            {

                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenHelperAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                var fileContentPostBody = new StringContent(fileContent, System.Text.Encoding.UTF8, "text/plain");

                // Endpoint for content in an existing file. Use "/me/drive/root/children/<file name>/content" if you know the name but not the Id.
                Uri fileEndpoint = new Uri(serviceEndpoint + "me/drive/items/" + fileId + "/content");

                var requestMessage = new HttpRequestMessage(HttpMethod.Put, fileEndpoint)
                {
                    Content = fileContentPostBody
                };


                HttpResponseMessage response = await client.SendAsync(requestMessage);

                if (response.IsSuccessStatusCode)
                {
                    fileUpdated = true;
                    Debug.WriteLine("Updated file Id: " + fileId);

                }
                else
                {
                    Debug.WriteLine("We could not update the file. The request returned this status code: " + response.StatusCode);
                    fileUpdated = false;
                }


            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not update the file. The request returned this status code: " + e.Message);
                fileUpdated = false;
            }

            return fileUpdated;
        }

        // Deletes a file in the user's root directory.
        public static async Task<bool> DeleteFileAsync(string fileId)
        {
            bool fileDeleted = false;

            try
            {

                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenHelperAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                // Endpoint for the file to delete.
                Uri fileEndpoint = new Uri(serviceEndpoint + "me/drive/items/" + fileId);

                HttpResponseMessage response = await client.DeleteAsync(fileEndpoint);

                if (response.IsSuccessStatusCode)
                {
                    fileDeleted = true;
                    Debug.WriteLine("Deleted file Id: " + fileId);

                }
                else
                {
                    Debug.WriteLine("We could not delete the file. The request returned this status code: " + response.StatusCode);
                    fileDeleted = false;
                }


            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not delete the file. The request returned this status code: " + e.Message);
                fileDeleted = false;
            }

            return fileDeleted;
        }

        // Copies a file in the user's root directory.
        public static async Task<bool> CopyFileAsync(string fileId, string copyFileName)
        {
            bool fileCopied = false;

            try
            {

                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenHelperAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                // Endpoint for the file to copy.
                Uri fileEndpoint = new Uri(serviceEndpoint + "me/drive/items/" + fileId + "/microsoft.graph.copy");

                // Build contents of post body and convert to StringContent object.
                // Using line breaks for readability.
                string postBody = "{'parentReference':{"
                    + "'path':'" + serviceEndpoint + "/drive/root:'},"
                    + "'name':'" + copyFileName + "'}";

                var copyBody = new StringContent(postBody, System.Text.Encoding.UTF8, "application/json");

                HttpResponseMessage response = await client.PostAsync(fileEndpoint, copyBody);

                if (response.IsSuccessStatusCode)
                {
                    fileCopied = true;
                    Debug.WriteLine("Copied file Id: " + fileId);

                }
                else
                {
                    Debug.WriteLine("We could not copy the file. The request returned this status code: " + response.StatusCode);
                    fileCopied = false;
                }


            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not copy the file. The request returned this status code: " + e.Message);
                fileCopied = false;
            }

            return fileCopied;
        }

        // Renames a file in the user's root directory.
        public static async Task<bool> RenameFileAsync(string fileId, string newFileName)
        {
            bool fileCopied = false;

            try
            {

                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenHelperAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                // Endpoint for the file to rename.
                Uri fileEndpoint = new Uri(serviceEndpoint + "me/drive/items/" + fileId);

                // Build contents of post body and convert to StringContent object.
                // Using line breaks for readability.
                string patchBody = "{"
                    + "'name':'" + newFileName + "'}";

                var copyBody = new StringContent(patchBody, System.Text.Encoding.UTF8, "application/json");

                var method = new HttpMethod("PATCH");

                var requestMessage = new HttpRequestMessage(method, fileEndpoint)
                {
                    Content = copyBody
                };

                HttpResponseMessage response = await client.SendAsync(requestMessage);

                if (response.IsSuccessStatusCode)
                {
                    fileCopied = true;
                    Debug.WriteLine("Renamed file Id: " + fileId);

                }
                else
                {
                    Debug.WriteLine("We could not rename the file. The request returned this status code: " + response.StatusCode);
                    fileCopied = false;
                }


            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not rename the file. The request returned this status code: " + e.Message);
                fileCopied = false;
            }

            return fileCopied;
        }


        // Creates a folder in the user's root directory.
        public static async Task<string> CreateFolderAsync(string folderName)
        {
            string createFolderId = null;
            JObject jResult = null;

            try
            {

                HttpClient client = new HttpClient();
                var token = await AuthenticationHelper.GetTokenHelperAsync();
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);

                // Endpoint for all files and folders for a user
                Uri foldersEndpoint = new Uri(serviceEndpoint + "me/drive/root/children");

                var folderMetadata = "{"
                    + "'name': '" + folderName + "',"
                    + "'folder': {},"
                    + "'@name.conflictBehavior': 'rename'"
                    + "}"
                    ;


                var folderMetadataPostBody = new StringContent(folderMetadata, System.Text.Encoding.UTF8, "application/json");

                HttpResponseMessage response = await client.PostAsync(foldersEndpoint, folderMetadataPostBody);

                if (response.IsSuccessStatusCode)
                {
                    string responseContent = await response.Content.ReadAsStringAsync();
                    jResult = JObject.Parse(responseContent);
                    createFolderId = (string)jResult["id"];
                    Debug.WriteLine("Created folder Id: " + createFolderId);


                }
                else
                {
                    Debug.WriteLine("We could not create the folder. The request returned this status code: " + response.StatusCode);
                    return null;
                }


            }

            catch (Exception e)
            {
                Debug.WriteLine("We could not create the folder. The request returned this status code: " + e.Message);
                return null;
            }

            return createFolderId;
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