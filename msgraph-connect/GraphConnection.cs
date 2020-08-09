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
using System.Net.Http;
using System.Net.Http.Headers;
using Microsoft.Graph;
using Microsoft.Identity.Client;

namespace msgraph_connect
{
    public class GraphHttpException : Exception
    {
        public GraphHttpException(int statusCode, string message, HttpResponseHeaders headers) : base(message) {
            this.StatusCode = statusCode;
            this.Headers = headers;
        }

        public int StatusCode {
            get;
            private set;
        }

        public HttpResponseHeaders Headers {
            get;
            private set;
        }
    }

    public class GraphConnection
    {
        private static string defaultAppId = "53316905-a6e5-46ed-b0c9-524a2379579e";
        private static string defaultGraphHost = "graph.microsoft.com";
        private static string defaultGraphUri = $"https://{defaultGraphHost}";

        private static string defaultLoginHost = "login.microsoftonline.com";
        private static string defaultLoginUri = $"{defaultLoginHost}";

        private Uri graphUri;
        private string apiVersion;
        private Uri loginAuthority;
        private IEnumerable<string> scopedPermissions;
        private Guid appId;

        private IPublicClientApplication app;
        private HttpClient graphClient;

        public GraphConnection(string[] permissions = null, string graphUri = "https://graph.microsoft.com", string loginUri = "https://login.microsoftonline.com", string apiVersion = "v1.0", string appId = null)
        {
            this.graphUri = new Uri(graphUri);
            this.apiVersion = apiVersion;
            this.loginAuthority = GetLoginAuthority(new Uri(loginUri));
            this.scopedPermissions = GetScopedPermissions(this.graphUri, permissions);

            var targetAppId = appId;

            if ( targetAppId == null )
            {
                targetAppId = GraphConnection.defaultAppId;
            }

            this.appId = new Guid(targetAppId);
        }

        public void Connect()
        {
            if ( this.app == null ) {
                var app = PublicClientApplicationBuilder.Create(this.appId.ToString()).WithAuthority(this.loginAuthority.ToString()).Build();

                var authenticationProvider = new DelegateAuthenticationProvider(
                    (requestMessage) => {
                        var authenticationResult = GetAccessToken();

                        requestMessage
                        .Headers
                        .Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", authenticationResult.AccessToken);

                        return Task.FromResult(0);
                    });

                this.graphClient = GraphClientFactory.Create(authenticationProvider, this.apiVersion, GraphClientFactory.Global_Cloud, null, null);
                this.app = app;
            }
        }

        public void Disconnect()
        {
            this.app = null;
            this.graphClient = null;
        }

        public AuthenticationResult GetAccessToken()
        {
            var task = GetAccessTokenAsync();

            return task.Result;
        }

        public async Task<AuthenticationResult> GetAccessTokenAsync()
        {
            var accounts = await this.app.GetAccountsAsync();

            IAccount existingAccount = null;

            var enumerator = accounts.GetEnumerator();
            if ( enumerator.MoveNext() )
            {
                existingAccount = enumerator.Current;
            }

            AuthenticationResult result;

            try
            {
                result = await this.app.AcquireTokenSilent(this.scopedPermissions, existingAccount)
                    .ExecuteAsync();
            }
            catch ( MsalUiRequiredException )
            {
                result = await this.app.AcquireTokenWithDeviceCode(this.scopedPermissions,
                                                                   deviceCodeResult =>
                    {
                        Console.WriteLine(deviceCodeResult.Message);
                        return Task.FromResult(0);
                    }).ExecuteAsync();
            }

            return result;
        }

        public async Task<HttpResponseMessage> InvokeRequestAsync(string relativeUri, string method = "GET")
        {
            Connect();

            var request = new HttpRequestMessage(new HttpMethod(method), relativeUri);

            return await this.graphClient.SendAsync(request);
        }

        public HttpResponseMessage InvokeRequest(string relativeUri, string method = "GET")
        {
            var task = InvokeRequestAsync(relativeUri, method);
            return task.Result;
        }

        public string InvokeRequestAndDeserialize(string relativeUri, string method = "GET")
        {
            var response = InvokeRequest(relativeUri, method);
            var task = response.Content.ReadAsStringAsync();
            var content = task.Result;

            if ( response.IsSuccessStatusCode ) {
                return content;
            } else {
                throw new GraphHttpException((int) response.StatusCode, content, response.Headers);
            }
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

                if ( loginUri.Host != GraphConnection.defaultLoginHost )
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

            // scopedPermissions.Add(".default", ".default");

            bool isDefaultGraphHost = IsPublicGraphHost( resourceUri );

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

        private bool IsPublicGraphHost(Uri graphUri)
        {
            return graphUri.Host == GraphConnection.defaultGraphHost;
        }

    }
}
