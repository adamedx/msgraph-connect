// Copyright 2020, Adam Edwards
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Identity.Client;

namespace msgraph_connect
{
    class GraphApplication
    {
        private static string defaultAppId = "53316905-a6e5-46ed-b0c9-524a2379579e";

        private static string defaultGraphHost = "graph.microsoft.com";
        private static string defaultGraphUri = $"https://{defaultGraphHost}";

        private static string defaultLoginHost = "login.microsoftonline.com";
        private static string defaultLoginUri = $"{defaultLoginHost}";

        private Uri graphUri;
        private Uri loginAuthority;
        private Guid appId;

        private IPublicClientApplication app;

        public GraphApplication(Uri graphUri, Uri loginUri, string appId = null)
        {
            this.graphUri = graphUri;

            if ( this.graphUri == null )
            {
                this.graphUri = new Uri(GraphApplication.defaultGraphUri);
            }

            this.loginAuthority = GetLoginAuthority(loginUri);

            var targetAppId = appId;

            if ( targetAppId == null )
            {
                targetAppId = GraphApplication.defaultAppId;
            }

            this.appId = new Guid(targetAppId);
        }

        public AuthenticationResult GetAccessToken(string[] permissions)
        {
            var task = GetAccessTokenAsync(permissions);

            return task.Result;
        }

        public async Task<AuthenticationResult> GetAccessTokenAsync(string[] permissions)
        {
            if ( this.app == null )
            {
                this.app = PublicClientApplicationBuilder.Create(this.appId.ToString()).WithAuthority(this.loginAuthority.ToString()).Build();
            }

            var existingAccount = await GetExistingAccountAsync();
            var scopedPermissions = GetScopedPermissions(this.graphUri, permissions);

            AuthenticationResult result;

            try
            {
                result = await this.app.AcquireTokenSilent(scopedPermissions, existingAccount)
                    .ExecuteAsync();
            }
            catch (MsalUiRequiredException)
            {
                result = await this.app.AcquireTokenWithDeviceCode(
                    scopedPermissions,
                    deviceCodeResult =>
                    {
                        Console.WriteLine(deviceCodeResult.Message);
                        return Task.FromResult(0);
                    }).ExecuteAsync();
            }

            return result;
        }

        private Uri GetLoginAuthority(Uri loginUri)
        {
            var tenantScope = "";
            var suffix = "";

            if ( ( loginUri.Segments != null ) &&
                 ( ( loginUri.Segments.Length < 2 ) ||
                   ( loginUri.Segments.Length == 2 && loginUri.Segments[0] == "/" ) ) )
            {
                suffix = "/oauth2/v2.0";

                tenantScope = "/common";

                if ( ! IsDefaultLoginHost(loginUri) )
                {
                    tenantScope = "/organizations";
                }
            }

            var authorityUri = new Uri($"https://{loginUri.Host}{tenantScope}{suffix}");

            return authorityUri;
        }

        private IEnumerable<string> GetScopedPermissions(Uri resourceUri, string[] permissions = null)
        {
            var scopedPermissions = new SortedList<string,string>();

            bool isDefaultGraphHost = IsDefaultGraphHost(resourceUri);

            string resourceString = resourceUri.ToString().TrimEnd('/');

            if ( permissions != null )
            {
                foreach ( var permission in permissions )
                {
                    var scopedPermission = permission;

                    switch ( permission ) {
                    case "openid":
                    case "profile":
                    case "offline_access":
                        continue;
                    }

                    if ( isDefaultGraphHost && ! scopedPermissions.ContainsKey( permission ) )
                    {
                        scopedPermission = $"{resourceString}/{permission}";
                    }

                    scopedPermissions.Add(permission, scopedPermission);
                }
            }

            return scopedPermissions.Values;
        }

        private bool IsDefaultGraphHost(Uri graphUri)
        {
            return graphUri.Host == GraphApplication.defaultGraphHost;
        }

        private bool IsDefaultLoginHost(Uri loginUri)
        {
            return loginUri.Host == GraphApplication.defaultLoginHost;
        }

        private async Task<IAccount> GetExistingAccountAsync()
        {
            var accounts = await this.app.GetAccountsAsync();

            IAccount existingAccount = null;

            var enumerator = accounts.GetEnumerator();
            if ( enumerator.MoveNext() )
            {
                existingAccount = enumerator.Current;
            }

            return existingAccount;
        }
    }
}
