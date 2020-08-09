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
    public class GraphConnection
    {
        private GraphApplication app;
        private HttpClient graphClient;
        private string appId;

        public Uri GraphUri
        {
            get;
            private set;
        }

        public string ApiVersion
        {
            get;
            private set;
        }

        public Uri LoginUri
        {
            get;
            private set;
        }

        public string[] Permissions
        {
            get;
            private set;
        }

        public string AppId
        {
            get
            {
                string appId = null;

                if ( this.app != null )
                {
                    appId = this.app.AppId.ToString();
                }

                return appId;
            }
        }

        public GraphConnection(string[] permissions = null, string graphUri = "https://graph.microsoft.com", string loginUri = "https://login.microsoftonline.com", string apiVersion = "v1.0", string appId = null)
        {
            this.Permissions = permissions;
            this.appId = appId;
            this.GraphUri = new Uri(graphUri);
            this.LoginUri = new Uri(loginUri);
            this.ApiVersion = apiVersion;
        }

        public void Connect()
        {
            if ( this.app == null )
            {
                this.app = new GraphApplication(this.GraphUri, this.LoginUri, this.appId);

                var authenticationProvider = new DelegateAuthenticationProvider(
                    (requestMessage) => {
                        var authenticationResult = this.app.GetAccessToken(this.Permissions);

                        requestMessage
                        .Headers
                        .Authorization = new AuthenticationHeaderValue("Bearer", authenticationResult.AccessToken);

                        return Task.FromResult(0);
                    });

                this.graphClient = GraphClientFactory.Create(authenticationProvider, this.ApiVersion, GraphClientFactory.Global_Cloud, null, null);
            }
        }

        public void Disconnect()
        {
            this.app = null;
            this.graphClient = null;
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
    }
}
