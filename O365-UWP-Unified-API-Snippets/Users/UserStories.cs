// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Windows.Storage;

namespace O365_UWP_Unified_API_Snippets
{
    class UserStories
    {
        private static readonly string STORY_DATA_IDENTIFIER = Guid.NewGuid().ToString();
        private static readonly string DEFAULT_MESSAGE_BODY = "This message was sent from the Office 365 UWP Unified API Snippets project";
        public static ApplicationDataContainer _settings = ApplicationData.Current.RoamingSettings;

        public static async Task<bool> TryGetMeAsync()
        {
            var currentUser = await UserSnippets.GetMeAsync();

            return currentUser != null;
        }

        public static async Task<bool> TryGetUsersAsync()
        {
            var users = await UserSnippets.GetUsersAsync();
            return users != null;
        }

        public static async Task<bool> TryCreateUserAsync()
        {
            string createdUser = await UserSnippets.CreateUserAsync(STORY_DATA_IDENTIFIER);
            return createdUser != null;
        }

        public static async Task<bool> TryGetCurrentUserDriveAsync()
        {
            string driveId = await UserSnippets.GetCurrentUserDriveAsync();
            return driveId != null;
        }

        public static async Task<bool> TryGetEventsAsync()
        {
            var events = await UserSnippets.GetEventsAsync();
            return events != null;
        }

        public static async Task<bool> TryCreateEventAsync()
        {
            string createdEvent = await UserSnippets.CreateEventAsync();
            return createdEvent != null;
        }

        public static async Task<bool> TryUpdateEventAsync()
        {
            // Create an event first, then update it.
            string createdEvent = await UserSnippets.CreateEventAsync();
            return await UserSnippets.UpdateEventAsync(createdEvent);
        }

        public static async Task<bool> TryDeleteEventAsync()
        {
            // Create an event first, then delete it.
            string createdEvent = await UserSnippets.CreateEventAsync();
            return await UserSnippets.DeleteEventAsync(createdEvent);
        }

        public static async Task<bool> TryGetMessages()
        {
            var messages = await UserSnippets.GetMessagesAsync();
            return messages != null;
        }

        public static async Task<bool> TrySendMailAsync()
        {
            return await UserSnippets.SendMessageAsync(
                    STORY_DATA_IDENTIFIER,
                    DEFAULT_MESSAGE_BODY,
                    (string)_settings.Values["userEmail"]
                );
        }

        public static async Task<bool> TryGetCurrentUserManagerAsync()
        {
            string managerName = await UserSnippets.GetCurrentUserManagerAsync();
            return managerName != null;
        }

        public static async Task<bool> TryGetDirectReportsAsync()
        {
            var users = await UserSnippets.GetDirectReportsAsync();
            return users != null;
        }

        public static async Task<bool> TryGetCurrentUserPhotoAsync()
        {
            string photoId = await UserSnippets.GetCurrentUserPhotoAsync();
            return photoId != null;
        }

        public static async Task<bool> TryGetCurrentUserGroupsAsync()
        {
            var groups = await UserSnippets.GetCurrentUserGroupsAsync();
            return groups != null;
        }

        public static async Task<bool> TryGetCurrentUserFilesAsync()
        {
            var files = await UserSnippets.GetCurrentUserFilesAsync();
            return files != null;
        }

        public static async Task<bool> TryCreateFileAsync()
        {
            var createdFileId = await UserSnippets.CreateFileAsync(Guid.NewGuid().ToString(), STORY_DATA_IDENTIFIER);
            return createdFileId != null;
        }

        public static async Task<bool> TryDownloadFileAsync()
        {
            var createdFileId = await UserSnippets.CreateFileAsync(Guid.NewGuid().ToString(), STORY_DATA_IDENTIFIER);
            var fileContent = await UserSnippets.DownloadFileAsync(createdFileId);
            return fileContent != null;
        }

        public static async Task<bool> TryUpdateFileAsync()
        {
            var createdFileId = await UserSnippets.CreateFileAsync(Guid.NewGuid().ToString(), STORY_DATA_IDENTIFIER);
            return await UserSnippets.UpdateFileAsync(createdFileId, STORY_DATA_IDENTIFIER);
        }

        public static async Task<bool> TryRenameFileAsync()
        {
            var newFileName = Guid.NewGuid().ToString();
            var createdFileId = await UserSnippets.CreateFileAsync(Guid.NewGuid().ToString(), STORY_DATA_IDENTIFIER);
            return await UserSnippets.RenameFileAsync(createdFileId, newFileName);
        }

        public static async Task<bool> TryDeleteFileAsync()
        {
            var fileName = Guid.NewGuid().ToString();
            var createdFileId = await UserSnippets.CreateFileAsync(Guid.NewGuid().ToString(), STORY_DATA_IDENTIFIER);
            return await UserSnippets.DeleteFileAsync(createdFileId);
        }

        public static async Task<bool> TryCreateFolderAsync()
        {
            var createdFolderId = await UserSnippets.CreateFolderAsync(Guid.NewGuid().ToString());
            return createdFolderId != null;
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