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
using System.Net.Http.Headers;

namespace msgraph_connect
{
    public class GraphHttpException : Exception
    {
        public GraphHttpException(int statusCode, string message, HttpResponseHeaders headers) : base(message)
        {
            this.StatusCode = statusCode;
            this.Headers = headers;
        }

        public int StatusCode
        {
            get;
            private set;
        }

        public HttpResponseHeaders Headers
        {
            get;
            private set;
        }
    }
}
