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
using msgraph_connect;

namespace graphcom
{
    class Program
    {
        static int Main(string[] args)
        {
            var connection = new GraphConnection();

            if ( args.Length < 1 )
            {
                PrintUsage();

                return 1;
            }

            string responseContent;

            try
            {
                responseContent = connection.InvokeRequestAndDeserialize(args[0]);
            }
            catch ( GraphHttpException exception )
            {
                Console.Error.WriteLine("Request failed with status code {0}", exception.StatusCode);
                Console.Error.WriteLine(exception.Message);
                return 1;
            }

            Console.WriteLine(responseContent);

            return 0;
        }

        static void PrintUsage()
        {
            Console.WriteLine("Usage:\n");
            Console.WriteLine("\tgraphcom <graphuri>\n");
        }
    }
}
